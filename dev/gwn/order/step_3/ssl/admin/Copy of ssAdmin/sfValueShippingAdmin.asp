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

Class clssfValueShipping
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private plngvalID
Private pdblvalShpAmt
Private pdblvalShpPurTotal

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


Public Property Get valID()
    valID = plngvalID
End Property

Public Property Get valShpAmt()
    valShpAmt = pdblvalShpAmt
End Property

Public Property Get valShpPurTotal()
    valShpPurTotal = pdblvalShpPurTotal
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    plngvalID = trim(rs("valID"))
    pdblvalShpAmt = trim(rs("valShpAmt"))
    pdblvalShpPurTotal = trim(rs("valShpPurTotal"))

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
        plngvalID = Trim(.Item("valID"))
        pdblvalShpAmt = Trim(.Item("valShpAmt"))
        pdblvalShpPurTotal = Trim(.Item("valShpPurTotal"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Find(lngID)

'On Error Resume Next

    With pRS
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngID) <> 0 Then
                .Find "valID=" & lngID
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

    Set pRS = GetRS("Select * from sfValueShipping Order By valShpPurTotal")
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

'***********************************************************************************************

Public Function Delete(lngvalID)

Dim sql

'On Error Resume Next

    sql = "Delete from sfValueShipping where valID = " & lngvalID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "Shipping Charge successfully deleted."
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

'On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    strErrorMessage = ValidateValues
    If ValidateValues Then
        If Len(plngvalID) = 0 Then plngvalID = 0

        sql = "Select * from sfValueShipping where valID = " & plngvalID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("valShpAmt") = pdblvalShpAmt
        rs("valShpPurTotal") = pdblvalShpPurTotal

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        plngvalID = rs("valID")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = "The shipping charge was successfully added."
            Else
                pstrMessage = "The shipping charge was successfully updated."
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
Dim pstrURL
 
    Response.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    Response.Write "<COLGROUP align='center'>"
    Response.Write "<COLGROUP align='center'>"
    Response.Write "	<tr class='tblhdr'>"
    Response.Write "		<th>Order totals up to</th>"
    Response.Write "		<th>Shipping Charge</th>"
    Response.Write "	</tr>"
	Response.Write "<tr><td colspan=2>"
    Response.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	Response.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary' id='tblSummary' " _
				 & ">"
    Response.Write "<COLGROUP align='center'>"
    Response.Write "<COLGROUP align='center'>"

    If prs.RecordCount > 0 Then
        prs.MoveFirst
        For i = 1 To prs.RecordCount
			pstrURL = "sfValueShippingAdmin.asp?Action=View&valID=" & prs("valID")
			
            If Trim(prs("valID")) = plngvalID Then
                Response.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
            Else
                Response.Write " <TR title='Click to edit this entry' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
            End If
         	Response.Write "<TD><a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to edit this entry'>" & FormatCurrency(prs("valShpPurTotal"),2) & "</a></TD>"
'           Response.Write "<TD>" & FormatCurrency(prs("valShpPurTotal"),2) & "</TD>" & vbCrLf
            Response.Write "<TD>" &  FormatCurrency(prs("valShpAmt"),2) & "</TD></TR>" & vbCrLf
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

    If Not IsNumeric(pdblvalShpAmt) And Len(pdblvalShpAmt) <> 0 Then
        strError = strError & "Please enter a number for the Shipping Charge." & cstrDelimeter
    ElseIf Len(pdblvalShpAmt) = 0 Then
        strError = strError & "Please enter a value for the Shipping Charge." & cstrDelimeter
    End If

    If Not IsNumeric(pdblvalShpPurTotal) And Len(pdblvalShpPurTotal) <> 0 Then
        strError = strError & "Please enter a number for the Order Total." & cstrDelimeter
    ElseIf Len(pdblvalShpPurTotal) = 0 Then
        strError = strError & "Please enter a value for the Order Total." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues
End Class   'clssfValueShipping

mstrPageTitle = "Value Based Shipping Charge Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfValueShipping

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclssfValueShipping = New clssfValueShipping
    
    Select Case mAction
        Case "New", "Update"
            mclssfValueShipping.Update
            If mclssfValueShipping.LoadAll Then mclssfValueShipping.Find Request.Form("valID")
        Case "Delete"
            mclssfValueShipping.Delete Request.Form("valID")
            mclssfValueShipping.LoadAll
        Case "View"
            If mclssfValueShipping.LoadAll Then mclssfValueShipping.Find Request.QueryString("valID")
        Case "Activate", "Deactivate"
            mclssfValueShipping.Activate Request.QueryString("valID"), mAction= Activate 
            If mclssfValueShipping.LoadAll Then mclssfValueShipping.Find Request.QueryString("valID")
        Case Else
            mclssfValueShipping.LoadAll
    End Select
    
Call WriteHeader("",True)
    With mclssfValueShipping
%>

<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--

function SetDefaults(theForm)
{
    theForm.valID.value = "";
    theForm.valShpAmt.value = "";
    theForm.valShpPurTotal.value = "";
return(true);
}

function btnNew_onclick(theButton)
{
var theForm = theButton.form;

    SetDefaults(theForm);
    theForm.btnUpdate.value = "Add New Shipping Rate";
    theForm.btnDelete.disabled = true;
}

function btnDelete_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete this shipping charge?");
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
    theForm.valID.value = "";
    theForm.valShpAmt.value = "";
    theForm.valShpPurTotal.value = "";

    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    theForm.btnDelete.disabled = false;
}

function ValidInput(theForm)
{
  if (!isNumeric(theForm.valShpPurTotal,false,"Please enter a value for the Order Total.")) {return(false);}
  if (!isNumeric(theForm.valShpAmt,false,"Please enter a value for the Shipping Charge.")) {return(false);}
  
  return(true);
}

//-->
</SCRIPT>

<BODY>
<CENTER>
<div class="pagetitle "><%= mstrPageTitle %></div>
<%= .OutputMessage %>
<%= .OutputSummary %>

<FORM action='sfValueShippingAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput(this);' method=post>
<input type=hidden id=valID name=valID value=<%= .valID %>>
<input type=hidden id=Action name=Action value='Update'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<colgroup align="right">
<colgroup align="left">
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spanDetailTitle">Edit Value Based Shipping Entry</span></th>
  </tr>

      <TR>
        <TD align=right>&nbsp;<LABEL id=lblvalShpPurTotal for=valShpPurTotal>Order totals up to:</LABEL>&nbsp;</TD>
        <TD>&nbsp;<INPUT id=valShpPurTotal name=valShpPurTotal Value='<%= .valShpPurTotal %>'></TD>
      </TR>
      <TR>
        <TD align=right>&nbsp;<LABEL id=lblvalShpAmt for=valShpAmt>Incur a shipping charge of:</LABEL>&nbsp;</TD>
        <TD>&nbsp;<INPUT id=valShpAmt name=valShpAmt Value='<%= .valShpAmt %>'></TD>
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
    Set mclssfValueShipping = Nothing
    Set cnn = Nothing
    Response.Flush
%>
