<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   1.00.0001												*
'*   Release Date:		December 26, 2003										*
'*   Release Date:		December 26, 2003										*
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

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	Dim maryQBTemplates(2)
	
	maryQBTemplates(0) = Array("Customers", "Save customer export to a local file", "qbCustomersExportIIF.xsl")
	maryQBTemplates(1) = Array("Sales Receipts", "Save sale receipts to a local file", "qbReceiptsExportIIF.xsl")
	maryQBTemplates(2) = Array("Invoices", "Save invoices to a local file", "qbInvoicesExportIIF.xsl")

'/
'/////////////////////////////////////////////////



'--------------------------------------------------------------------------------------------------
%>
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_class.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrdersToXML.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
mstrPageTitle = "Sandshot Software's Order Manager to QuickBooks Export Tool"
Call WriteHeader("",True)
%>
<SCRIPT language="vbscript">
dim maryCells()

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function saveFile(byRef strFileName, byRef strContents)
	'Initialize and script ActiveX controls not marked as safe needs to be set to Enable or Promopt

	Dim fso
	Dim MyFile
	Dim pstrErrorMessage
	
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
			Set MyFile = fso.CreateTextFile(strFileName, True)
			MyFile.Write strContents
			MyFile.close
			Set MyFile = Nothing
		End If

		Set fso = Nothing
		
		saveFile = (Err.number = 0)

	End Function	'saveFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

	Function saveExportFile(strExportSection)

	Dim pobjFileChooser
	Dim pobjExportSection
	Dim pstrFile
	Dim pstrTextToSave

		Set pobjExportSection = document.all(strExportSection)
		Set pobjFileChooser = document.all("tempFile")
		pobjFileChooser.click()

		pstrFile = pobjFileChooser.value
		If Len(pstrFile) > 0 Then
			pstrTextToSave = pobjExportSection.innerTEXT & vbcrlf	'line break added since innerTEXT will remove the trailing returns
			Call saveFile(pstrFile, pstrTextToSave)
		End If

		saveExportFile = False

		Set pobjFileChooser = Nothing
		Set pobjExportSection = Nothing
		
		saveExportFile = (Err.number <> 0)

	End Function	'saveExportFile

	'-----------------------------------------------------------------------------------------------------------------------------------------------

</SCRIPT>
<%
'**************************************************
'
'	Start Code Execution
'

'page variables
Dim mstrXSLFilePath
Dim pstrOrderIDs
Dim paryOrderIDs
Dim paryOrderIDsToExport
Dim paryInvoices
Dim mstrQBStartingInvoice
Dim mstrAction

	mstrAction = LoadRequestValue("Action")

    If Len(Request.Form("startingInvoice")) > 0 Then
		mstrQBStartingInvoice = Request.Form("startingInvoice")
	Else
		mstrQBStartingInvoice = mstrLastOrderExported
	End If

    pstrOrderIDs = Request.Form("chkssOrderID")
	paryOrderIDs = Split(pstrOrderIDs, ", ")
	
	mstrXSLFilePath = ssAdminPath & "exportTemplates/qbTemplates/"
	
	If cblnAutoExport Then
		Call setExportedStatus(pstrOrderIDs, "ssExported", 1)
	Else
		Call setExportedStatus(Request.Form("chkssOrderIDExported"), "ssExported", 1)
	End If
	
	ReDim paryInvoices(UBound(paryOrderIDs))   
	For i = UBound(paryOrderIDs) To 0 Step -1
		If cblnUseCustomInvoiceNumber Or Len(Request.Form("startingInvoice")) > 0 Then
			paryInvoices(i) =  CStr(mstrQBStartingInvoice + UBound(paryOrderIDs) - i + 1)
		Else
			paryInvoices(i) =  paryOrderIDs(i)
		End If
	Next 'i
   
%>
<table border=0 cellpadding=2 cellspacing=0>
<tr><th>&nbsp;</th><th align=left>Quick Books Export Tool</th></tr>
<tr>
  <td rowspan=7 valign=top>
  <table border="1" cellpadding="2" cellspacing="0">
    <tr>
      <td nowrap>
		<% For i = 0 To UBound(maryQBTemplates) %>
		<a href="" onclick="saveExportFile('exportFile<%= i %>'); return false;" title="<%= maryQBTemplates(i)(1) %>"><%= maryQBTemplates(i)(0) %></a><br />
		<% Next 'i %>
		<hr>
		<a href="" onclick="<% For i = 0 To UBound(maryQBTemplates) %>saveExportFile('exportFile<%= i %>');<% Next 'i %> return false;" title="Save all files">Save All</a><br />
		<span id=spantempFile style="display:none"><input type=file id=tempFile name=tempFile size="20"></span>
      </td>
    </tr>
  </table>
 </td>
  <td>
    <form name="frmQuickBooks" id="frmQuickBooks" action="ssOrderAdmin_QuickBooksExport.asp" method="post">
    <input type="hidden" name="chkssOrderID" id="chkssOrderID" value="<%= pstrOrderIDs %>">
    <input type="hidden" name="Action" id="Action" value="">
    <table border="1" cellpadding="2" cellspacing="0">
		<tr><td>Orders in this export (<%= UBound(paryOrderIDs)+1 %>):</td><% For i = 0 To UBound(paryOrderIDs) %><td align=center><%= paryOrderIDs(i) %></td><% Next 'i %></tr>    
		<% If Not cblnAutoExport Then %>
		<tr><td>Mark exported&nbsp;<input type=checkbox name="chkCheckAll" id="chkCheckAll" value="0" onclick="checkAll(this.form.chkssOrderIDExported,this.checked); this.form.markExported.disabled=(!anyChecked(this.form.chkssOrderIDExported));"></td>
			<% For i = 0 To UBound(paryInvoices)%>
			<td align="center"><input type=checkbox name="chkssOrderIDExported" id="chkssOrderIDExported" value="<%= paryInvoices(i) %>" onclick=" this.form.markExported.disabled=(!anyChecked(this.form.chkssOrderIDExported));" <%= isChecked(cblnAutoExport) %>></td>
			<% Next 'i %>
		</tr>    
		<% End If %>
		<tr><td>Invoice numbers to be exported:</td><% For i = 0 To UBound(paryInvoices) %><td align="center"><%= cstrInvoiceOrderPrefix & paryInvoices(i) %></td><% Next 'i %></tr>    
		<% If cblnUseCustomInvoiceNumber Then %>
		<tr><td><label for="startingInvoice" title="Invoice number will start one greater than this value if the custom invoice selection is chosen">Last Invoice Number Exported:</label></td><td align="center" colspan="<%= UBound(paryOrderIDs)+1 %>"><input type="text" name="startingInvoice" id="startingInvoice" value="<%= mstrQBStartingInvoice %>"></td></tr>
		<tr><td>&nbsp;</td><td colspan="<%= UBound(paryOrderIDs)+1 %>"><input type="submit" name="submit" value="Run Extract Again">
		<% If Not cblnAutoExport Then %>
		<input type="submit" name="markExported" id="markExported" value="Mark Exported" onclick="if (!anyChecked(this.form.chkssOrderIDExported)){alert('Please check an order to mark as exported'); return false;} this.form.Action.value='MarkExported'; return true;" disabled>   
		<% End If %>
		</td></tr>
		<% End If	'cblnUseCustomInvoiceNumber %>
    </table>
    </form>
  </td>
</tr>
<% For i = 0 To UBound(maryQBTemplates) %>
<tr>
  <td><hr><b><a href="" onclick="saveExportFile('exportFile<%= i %>'); return false;" title="<%= maryQBTemplates(i)(1) %>"><%= maryQBTemplates(i)(0) %></a></b><hr></td>
<tr>
  <td valign=top><pre id="exportFile<%= i %>"><%
	mstrLastOrderExported = mstrQBStartingInvoice
	Response.Write exportOrders(pstrOrderIDs, mstrXSLFilePath & "\" & maryQBTemplates(i)(2))
  %></pre></td>
</tr>
<% Next 'i %>
</table>
</body>
</html>
<%

Call ReleaseObject(cnn)
Call saveStartingInvoiceNumber(mstrLastOrderExported)

%>
