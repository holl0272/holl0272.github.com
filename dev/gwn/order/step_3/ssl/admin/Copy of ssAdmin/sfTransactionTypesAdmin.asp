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

Class clssfTransactionTypes
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private plngtransID
Private pbyttransIsActive
Private pstrtransName
Private pstrtransType

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


Public Property Get transID()
    transID = plngtransID
End Property

Public Property Get transIsActive()
    transIsActive = pbyttransIsActive
End Property

Public Property Get transName()
    transName = pstrtransName
End Property

Public Property Get transType()
    transType = pstrtransType
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    plngtransID = trim(rs("transID"))
    pbyttransIsActive = trim(rs("transIsActive"))
    pstrtransName = trim(rs("transName"))
    pstrtransType = trim(rs("transType"))

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
        plngtransID = Trim(.Item("transID"))
        pbyttransIsActive = (uCase(.Item("transIsActive")) = "ON")
        pstrtransName = Trim(.Item("transName"))
        pstrtransType = Trim(.Item("transType"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "transID=" & lngID
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

    Set pRS = GetRS("Select * from sfTransactionTypes")
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(lngtransID)

Dim sql

'On Error Resume Next

    sql = "Delete from sfTransactionTypes where transID = " & lngtransID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Record successfully deleted."
        Delete = True
    Else
        pstrMessage = Err.Description
        Delete = False
    End If

End Function    'Delete

'***********************************************************************************************

Public Function Activate(lngID,blnActivate)

Dim sql

'On Error Resume Next

	if blnActivate then
		sql = "Update sfTransactionTypes Set transIsActive=1 where transID = " & lngID
    else
		sql = "Update sfTransactionTypes Set transIsActive=0 where transID = " & lngID
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        pstrMessage = "Record successfully updated."
        Activate = True
    Else
        pstrMessage = Err.Description
        Activate = False
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
        If Len(plngtransID) = 0 Then plngtransID = 0

        sql = "Select * from sfTransactionTypes where transID = " & plngtransID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("transIsActive") = (pbyttransIsActive * -1)
        rs("transName") = pstrtransName
        rs("transType") = pstrtransType

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        plngtransID = rs("transID")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrtransName & " was successfully added."
            Else
                pstrMessage = pstrtransName & " was successfully updated."
            End If
        Else
            pblnError = True
        End If
    Else
        pblnError = True
    End If

    Update = (not pblnError)
    Application.Contents.Remove("CreditCardArray")
	Application.Contents.Remove("PaymentTypesArray")
	
End Function    'Update

'***********************************************************************************************


Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim pstrURL
Dim pstrTitle

    Response.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none'>"
    Response.Write "<colgroup align='left' width='33%'>"
    Response.Write "<colgroup align='center' width='33%'>"
    Response.Write "<colgroup align='center' width='34%'>"
    Response.Write "	<tr class='tblhdr'>"
    Response.Write "		<th>&nbsp;&nbsp;Method</th>"
    Response.Write "		<th>Type&nbsp;&nbsp;&nbsp;&nbsp;</th>"
    Response.Write "		<th>Active&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>"
    Response.Write "	</tr>"
	Response.Write "<tr><td colspan=3>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary' id='tblSummary' " _
				 & ">"
    Response.Write "<colgroup align='left' width='33%'>"
    Response.Write "<colgroup align='center' width='33%'>"
    Response.Write "<colgroup align='center' width='34%'>"
    If prs.RecordCount > 0 Then
        prs.MoveFirst
        For i = 1 To prs.RecordCount
			pstrURL = "sfTransactionTypesAdmin.asp?Action=View&transID=" & prs("transID")
			pstrTitle = "Click to view " & prs("transName") & "."
			
            If Trim(prs("transID")) = plngtransID Then
                Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
            Else
				if cBool(pRS("transIsActive")) then
					Response.Write " <TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
				else
					Response.Write " <TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
        		end if
            End If
        	Response.Write "<TD>&nbsp;&nbsp;&nbsp;<a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & pRS("transName") & "&nbsp;</a></TD>"
            Response.Write "<TD>" & prs("transType") & "&nbsp;</TD>" & vbCrLf
			if cBool(pRS("transIsActive")) then
        		Response.Write "<TD><a href='sfTransactionTypesAdmin.asp?Action=Deactivate&transID=" & pRS("transID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & prs("transName") & ".'>Active</a></TD></TR>" & vbCrLf
			else
        		Response.Write "<TD><a href='sfTransactionTypesAdmin.asp?Action=Activate&transID=" & pRS("transID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & prs("transName") & ".'>Inactive</a></TD></TR>" & vbCrLf
        	end if
            prs.MoveNext
        Next
    Else
        Response.Write "<TR><TD><h3>There are no Records</h3></TD></TR>"
    End If
    Response.Write "</td></tr></TABLE></div>"
    Response.Write "</TABLE>"

End Sub      'OutputSummary

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues
End Class   'clssfTransactionTypes

mstrPageTitle = "Transaction Type Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   Connection: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfTransactionTypes

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclssfTransactionTypes = New clssfTransactionTypes
    
    Select Case mAction
        Case "New", "Update"
            mclssfTransactionTypes.Update
            If mclssfTransactionTypes.LoadAll Then mclssfTransactionTypes.Find Request.Form("transID")
        Case "Delete"
            mclssfTransactionTypes.Delete Request.Form("transID")
            mclssfTransactionTypes.LoadAll
        Case "View"
            If mclssfTransactionTypes.LoadAll Then mclssfTransactionTypes.Find Request.QueryString("transID")
        Case "Activate", "Deactivate"
            mclssfTransactionTypes.Activate Request.QueryString("transID"), mAction = "Activate" 
            If mclssfTransactionTypes.LoadAll Then mclssfTransactionTypes.Find Request.QueryString("transID")
        Case Else
            mclssfTransactionTypes.LoadAll
    End Select
    
	Call WriteHeader("",True)
    With mclssfTransactionTypes
%>

<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var strDetailTitle = "Edit <%= .transName %> Detail";

function SetDefaults(theForm)
{
    theForm.transID.value = "";
    theForm.transIsActive.checked = false;
    theForm.transName.value = "";
    theForm.transType.value = "";
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.btnUpdate.value = "Add Transaction Type";
    theForm.btnDelete.disabled = true;
    document.all("spanDetailTitle").innerHTML = theForm.btnUpdate.value;
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.transName.value + "?");
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
    theForm.transID.value = "";
    theForm.transIsActive.value = "";
    theForm.transName.value = "";
    theForm.transType.value = "";

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDelete.disabled = false;
    document.all("spanDetailTitle").innerHTML = strDetailTitle;
}

function ValidInput(theForm)
{

    return(true);
}

//-->
</SCRIPT>

<BODY>
<CENTER>
<div class="pagetitle "><%= mstrPageTitle %></div>
<%= .OutputMessage %>
<%= .OutputSummary %>

<FORM action='sfTransactionTypesAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=transID name=transID value=<%= .transID %>>
<input type=hidden id=Action name=Action value='Update'>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="right">
<colgroup align="left">
<tr class='tblhdr'>
<th colspan="2" align=center><span id="spanDetailTitle">Edit <%= .transName %> Detail</span></th>
</tr>
      <TR>
        <TD class="label">Method:&nbsp;</TD>
        <TD>&nbsp;<INPUT id=transName name=transName Value='<%= .transName %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD class="label">Type&nbsp;</TD>
        <TD>&nbsp;
		<SELECT id=transType name=transType>
			<OPTION <% If .transType="Credit Card" Then Response.Write "Selected" %>>Credit Card</Option>
			<OPTION <% If .transType="eCheck" Then Response.Write "Selected" %>>eCheck</Option>
			<OPTION <% If .transType="COD" Then Response.Write "Selected" %>>COD</Option>
			<OPTION <% If .transType="PO" Then Response.Write "Selected" %>>PO</Option>
			<OPTION <% If .transType="PhoneFax" Then Response.Write "Selected" %>>PhoneFax</Option>
			<OPTION <% If .transType="InternetCash" Then Response.Write "Selected" %>>InternetCash</Option>
			<OPTION <% If .transType="PayPal" Then Response.Write "Selected" %>>PayPal</Option>
		</SELECT>
		</TD>
<!--        <TD>&nbsp;<INPUT id=transType name=transType Value='<%= .transType %>' maxlength=30 size=30></TD> -->
      </TR>
      <TR>
        <TD>&nbsp;</TD>
        <TD>&nbsp;<INPUT type=checkbox id=transIsActive name=transIsActive <% If (.transIsActive=1) then Response.Write "checked" %>>&nbsp;<LABEL id=lbltransIsActive for=transIsActive>Is Active</LABEL></TD>
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
    Set mclssfTransactionTypes = Nothing
    Set cnn = Nothing
    Response.Flush
%>
