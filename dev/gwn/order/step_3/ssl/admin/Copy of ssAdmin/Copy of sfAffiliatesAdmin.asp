<%Option Explicit
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version:	2.00.001		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		August 18, 2003											*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

Class clssfAffiliates
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private pstraffAddress1
Private pstraffAddress2
Private pstraffCity
Private pstraffCountry
Private pstraffEmail
Private pstraffFAX
Private pstraffHttpAddr
Private plngaffID
Private pstraffName
Private pstraffCompany
Private pstraffNotes
Private pstraffPhone
Private pstraffState
Private pstraffZip

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


Public Property Get affAddress1()
    affAddress1 = pstraffAddress1
End Property

Public Property Get affAddress2()
    affAddress2 = pstraffAddress2
End Property

Public Property Get affCity()
    affCity = pstraffCity
End Property

Public Property Get affCountry()
    affCountry = pstraffCountry
End Property

Public Property Get affEmail()
    affEmail = pstraffEmail
End Property

Public Property Get affFAX()
    affFAX = pstraffFAX
End Property

Public Property Get affHttpAddr()
    affHttpAddr = pstraffHttpAddr
End Property

Public Property Get affID()
    affID = plngaffID
End Property

Public Property Get affName()
    affName = pstraffName
End Property

Public Property Get affCompany()
    affCompany = pstraffCompany
End Property

Public Property Get affNotes()
    affNotes = pstraffNotes
End Property

Public Property Get affPhone()
    affPhone = pstraffPhone
End Property

Public Property Get affState()
    affState = pstraffState
End Property

Public Property Get affZip()
    affZip = pstraffZip
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    pstraffAddress1 = trim(rs("affAddress1"))
    pstraffAddress2 = trim(rs("affAddress2"))
    pstraffCity = trim(rs("affCity"))
    pstraffCountry = trim(rs("affCountry"))
    pstraffEmail = trim(rs("affEmail"))
    pstraffFAX = trim(rs("affFAX"))
    pstraffHttpAddr = trim(rs("affHttpAddr"))
    plngaffID = trim(rs("affID"))
    pstraffName = trim(rs("affName"))
    pstraffCompany = trim(rs("affCompany"))
    pstraffNotes = trim(rs("affNotes"))
    pstraffPhone = trim(rs("affPhone"))
    pstraffState = trim(rs("affState"))
    pstraffZip = trim(rs("affZip"))

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
        pstraffAddress1 = Trim(.Item("affAddress1"))
        pstraffAddress2 = Trim(.Item("affAddress2"))
        pstraffCity = Trim(.Item("affCity"))
        pstraffCountry = Trim(.Item("affCountry"))
        pstraffEmail = Trim(.Item("affEmail"))
        pstraffFAX = Trim(.Item("affFAX"))
        pstraffHttpAddr = Trim(.Item("affHttpAddr"))
        plngaffID = Trim(.Item("affID"))
        pstraffName = Trim(.Item("affName"))
        pstraffCompany = Trim(.Item("affCompany"))
        pstraffNotes = Trim(.Item("affNotes"))
        pstraffPhone = Trim(.Item("affPhone"))
        pstraffState = Trim(.Item("affState"))
        pstraffZip = Trim(.Item("affZip"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "affID=" & lngID
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues (pRS)
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function LoadAll()

'On Error Resume Next

    Set pRS = GetRS("Select * from sfAffiliates " & mstrsqlWhere)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(lngaffID)

Dim sql

'On Error Resume Next

    sql = "Delete from sfAffiliates where affID = " & lngaffID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "The Affiliate was successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Update()

Dim sql
Dim rs
Dim strErrorMessage
Dim blnAdd

On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If ValidateValues Then
        If Len(plngaffID) = 0 Then plngaffID = 0

        sql = "Select * from sfAffiliates where affID = " & plngaffID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("affAddress1") = pstraffAddress1
        rs("affAddress2") = pstraffAddress2
        rs("affCity") = pstraffCity
        rs("affCountry") = pstraffCountry
        rs("affEmail") = pstraffEmail
        rs("affFAX") = pstraffFAX
        rs("affHttpAddr") = pstraffHttpAddr
        rs("affName") = pstraffName
        rs("affCompany") = pstraffCompany
        rs("affNotes") = pstraffNotes
        rs("affPhone") = pstraffPhone
        rs("affState") = pstraffState
        rs("affZip") = pstraffZip

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        plngaffID = rs("affID")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstraffName & " was successfully added."
            Else
                pstrMessage = "The changes to " & pstraffName & " were successfully saved."
            End If
        Else
            pblnError = True
        End If
    Else
        pblnError = True
    End If

    Update = (not pblnError)

End Function    'Update

'***********************************************************************************************


Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim aSortHeader(3,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr

	With Response

    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='left' width='40%'>"
    .Write "<COLGROUP align='center' width='40%'>"
    .Write "<COLGROUP align='center' width='20%'>"
    .Write "  <tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort Affiliates in descending order"
		aSortHeader(2,0) = "Sort Emails in descending order"
		aSortHeader(3,0) = "Sort Web Sites in descending order"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort Manufacturers in ascending order"
		aSortHeader(2,0) = "Sort Emails in ascending order"
		aSortHeader(3,0) = "Sort Web Sites in ascending order"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "&nbsp;&nbsp;Affiliate"
	aSortHeader(2,1) = "Email"
	aSortHeader(3,1) = "Web Site"

	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	for i = 1 to 3
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
	next 'i

	Response.Write "<tr><td colspan=3>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' " _
				 & ">"
    .Write "<COLGROUP align='left' width='40%'>"
    .Write "<COLGROUP align='center' width='40%'>"
    .Write "<COLGROUP align='center' width='20%'>"

	If prs.RecordCount > 0 Then
        prs.MoveFirst
        For i = 1 To prs.RecordCount
			pstrAbbr = Trim(prs("affID"))
 			pstrTitle = "Click to edit " & prs("affName")
			pstrURL = "sfAffiliatesAdmin.asp?Action=View&affID=" & pstrAbbr

			if pstrAbbr = plngaffID then
        		.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
				.Write "<TD>" & prs("affName") & "</TD>" & vbCrLf
			else
				.Write "<TR title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
				.Write "<TD><a href='" & pstrURL & "'  onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & prs("affName") & "</a></TD>" & vbCrLf
        	end if

			If (len(prs("affEmail"))=0 or isNull(prs("affEmail"))) Then
				.Write "<TD>&nbsp;</TD>" & vbCrLf
			Else
				.Write "<TD><a href='mailto:" & prs("affEmail") & "'  onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Send " & prs("affName") & " an email.'>" & prs("affEmail") & "</a></TD>" & vbCrLf
			End If
			If (len(prs("affHttpAddr"))=0 or isNull(prs("affHttpAddr"))) Then
				.Write "<TD>&nbsp;</TD>" & vbCrLf
			Else
				.Write "<TD><a href='" & prs("affHttpAddr") & "'>" & prs("affHttpAddr") & "</a></TD>" & vbCrLf
			End If
            Response.Write "</TR>" & vbCrLf
            prs.MoveNext
        Next
    Else
        Response.Write "<TR><TD colspan=3 align=center><h3>There are no Affiliates</h3></TD></TR>"
    End If
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
	End With
	
End Sub      'OutputSummary

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues
End Class   'clssfAffiliates


Sub LoadFilter

dim pstrOrderBy

	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")
	blnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")

'Build Filter

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'Affiliate Name
			pstrOrderBy = "affName"
		Case "2"	'Affiliate email
			pstrOrderBy = "affEmail"
		Case "3"	'Affiliate WebSite
			pstrOrderBy = "affHttpAddr"
	End Select	

	If len(pstrOrderBy) > 0 then
		mstrsqlWhere = mstrsqlWhere & " Order By " & pstrOrderBy & " " & mstrSortOrder
	Else
		mstrsqlWhere = ""
	End If
	
End Sub    'LoadFilter

mstrPageTitle = "Affiliate Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfAffiliates
Dim mstrOrderBy, mstrSortOrder, blnShowSummary
Dim mvntID
Dim mstrsqlWhere

	Call LoadFilter
    mvntID = Request.QueryString("affID")
    If len(mvntID) = 0 Then mvntID = Request.Form("affID")

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclssfAffiliates = New clssfAffiliates
    
    Select Case mAction
        Case "New", "Update"
            mclssfAffiliates.Update
            If mclssfAffiliates.LoadAll Then mclssfAffiliates.Find mvntID
        Case "Delete"
            mclssfAffiliates.Delete mvntID
            mclssfAffiliates.LoadAll
        Case "View"
            If mclssfAffiliates.LoadAll Then mclssfAffiliates.Find mvntID
        Case "Activate", "Deactivate"
            mclssfAffiliates.Activate Request.QueryString("affID"), mAction= Activate 
            If mclssfAffiliates.LoadAll Then mclssfAffiliates.Find mvntID
        Case Else
            mclssfAffiliates.LoadAll
    End Select
    
	Call WriteHeader("body_onload();",True)
    With mclssfAffiliates
%>

<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var theKeyField;
var strDetailTitle = "<%= .affName %> Details";

function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = theDataForm.affID;
<% If blnShowSummary then Response.Write "DisplaySummary();" & vbcrlf %>
}

function SetDefaults(theForm)
{
    theForm.affAddress1.value = "";
    theForm.affAddress2.value = "";
    theForm.affCity.value = "";
    theForm.affCountry.value = "";
    theForm.affEmail.value = "";
    theForm.affFAX.value = "";
    theForm.affHttpAddr.value = "";
    theForm.affID.value = "";
    theForm.affName.value = "";
    theForm.affNotes.value = "";
    theForm.affPhone.value = "";
    theForm.affState.value = "";
    theForm.affZip.value = "";
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.btnUpdate.value = "Add Affiliate";
    theForm.btnDelete.disabled = true;
    theForm.affName.focus();
    document.all("spanDetailTitle").innerHTML = theDataForm.btnUpdate.value;
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.affName.value + "?");
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

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDelete.disabled = false;
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
}

function SortColumn(strColumn,strSortOrder)
{
	theDataForm.Action.value = "";
	theDataForm.OrderBy.value = strColumn;
	theDataForm.SortOrder.value = strSortOrder;
	theDataForm.submit();
	return false;
}

function ViewDetail(theValue)
{
	theKeyField.value = theValue;
	theDataForm.Action.value = "View";
	theDataForm.submit();
	return false;
}

function ValidInput(theForm)
{

    return(true);
}

//-->
</SCRIPT>

<BODY onload="body_onload();">
<CENTER>
<TABLE border=0 cellPadding=5 cellSpacing=1 width="95%">
  <TR>
    <TH><div class="pagetitle "><%= mstrPageTitle %></div></TH>
    <TH>&nbsp;</TH>
    <TH align='right'>
		<a href="#"><div id="divSummary" onclick="return DisplaySummary();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Summary">Hide Summary</div></a>
	</TH>
  </TR>
</TABLE>
<%= .OutputMessage %>

<FORM action='sfAffiliatesAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=affID name=affID value=<%= .affID %>>
<input type=hidden id=Action name=Action value='Update'>
<input type=hidden id=blnShowSummary name=blnShowSummary value=''>
<input type=hidden id=OrderBy name=OrderBy value='<%= mstrOrderBy %>'>
<input type=hidden id=SortOrder name=SortOrder value='<%= mstrSortOrder %>'>

<%= .OutputSummary %>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <colgroup align=right>
  <colgroup align=left>
  <tr class='tblhdr'>
	<th colspan="3" align=center><span id="spanDetailTitle"><%= .affName %> Details</span></th>
  </tr>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffName for=affName>Name:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affName name=affName Value='<%= .affName %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffCompany for=affName>Company:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affCompany name=affCompany Value='<%= .affCompany %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffAddress1 for=affAddress1>Address 1:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affAddress1 name=affAddress1 Value='<%= .affAddress1 %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffAddress2 for=affAddress2>Address 2:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affAddress2 name=affAddress2 Value='<%= .affAddress2 %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffCity for=affCity>City:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affCity name=affCity Value='<%= .affCity %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffState for=affState>State:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affState name=affState Value='<%= .affState %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffZip for=affZip>ZIP:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affZip name=affZip Value='<%= .affZip %>' maxlength=15 size=15></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffCountry for=affCountry>Country:</LABEL></TD>
        <TD>&nbsp;
			<select size="1"  id=affCountry name=affCountry>
			<% Call MakeCombo("Select loclctryName from sfLocalesCountry","loclctryName","loclctryName",.affCountry) %>
			</select>
		</TD>        
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffEmail for=affEmail>Email:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affEmail name=affEmail Value='<%= .affEmail %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffPhone for=affPhone>Phone:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affPhone name=affPhone Value='<%= .affPhone %>' maxlength=20 size=20></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffFAX for=affFAX>FAX:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affFAX name=affFAX Value='<%= .affFAX %>' maxlength=20 size=20></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffHttpAddr for=affHttpAddr>URL:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affHttpAddr name=affHttpAddr Value='<%= .affHttpAddr %>' maxlength=100 size=60></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblaffNotes for=affNotes>Notes:</LABEL></TD>
        <TD>&nbsp;<INPUT id=affNotes name=affNotes Value='<%= .affNotes %>' maxlength=255 size=60></TD>
      </TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick='return btnNew_onclick(this)'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)'>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDelete name=btnDelete type=button value=Delete onclick='return btnDelete_onclick(this)'>
        <INPUT class='butn' id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
  </TR>
</TABLE>
</FORM>

</CENTER>
</BODY>
</HTML>
<%
    End With
    Set mclssfAffiliates = Nothing
    Set cnn = Nothing
    Response.Flush
%>
