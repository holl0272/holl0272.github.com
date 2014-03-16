<% Option Explicit 
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

Class clsCoupon
'Assumptions:
'	Connection: defines a previously opened connection to the database
'

'database variables
dim pstrcpCouponCode
dim pbytcpActivate
dim pbytcpNeverExpire
dim pstrcpDescription
dim pdblcpValue
dim pstrcpType
dim pdtcpExpirationDate
dim pdblcpMin

'working variables
dim pConnection
dim prsCoupon
dim pstrMessage
dim pblnError

'variable for  handling Order parameters
dim aProductsInCart
dim pstrsfCoupons
dim psubTotal
dim pBestcpValueAmount
dim pBestcpValueCode

Public Property Let Connection(objConnection)
	Set pConnection = objConnection
End Property

Public Property Get Message
	Message = pstrMessage
End Property

Public Sub OutputMessage

dim i
dim aError

	aError = split(pstrMessage,";")
	for i = 0 to ubound(aError)
		if pblnError then
			Response.Write "<P align='center'><H4><FONT color=Red>" & aError(i) & "</FONT></H4></P>"	
		else
			Response.Write "<P align='center'><H4>" & aError(i) & "</H4></P>"	
		end if
	next 'i

End Sub	'OutputMessage

Public Property Let cpExpirationDate(dtcpExpirationDate)
	pdtcpExpirationDate = cpExpirationDate
End Property

Public Property Get cpExpirationDate
	cpExpirationDate = pdtcpExpirationDate
End Property

Public Property Let cpCouponCode(strcpCouponCode)
	pstrcpCouponCode = strcpCouponCode
End Property

Public Property Get cpCouponCode
	cpCouponCode = pstrcpCouponCode
End Property

Public Property Let cpValue(lngcpValue)
	pdblcpValue = cpValue
End Property

Public Property Get cpValue
	cpValue = pdblcpValue
End Property

Public Property Let cpType(blncpType)
	pstrcpType = cpType
End Property

Public Property Get cpType
	cpType = pstrcpType
End Property

Public Property Let cpMin(curSubTotal)
	pdblcpMin = cpMin
End Property

Public Property Get cpMin
	cpMin = pdblcpMin
End Property

Public Property Let cpDescription(strcpDescription)
	pstrcpDescription = cpDescription
End Property

Public Property Get cpDescription
	cpDescription = pstrcpDescription
End Property

Public Property Let cpActivate(blncpActivate)
	pbytcpActivate = blncpActivate
End Property

Public Property Get cpActivate
	cpActivate = pbytcpActivate
End Property

Public Property Let cpNeverExpire(blncpNeverExpire)
	pbytcpNeverExpire = blncpNeverExpire
End Property

Public Property Get cpNeverExpire
	cpNeverExpire = pbytcpNeverExpire
End Property

'***********************************************************************************************

Public Function FindBycpCouponCode(strcpCouponCode)

	with prsCoupon
		.Find "cpCouponCode = '" & strcpCouponCode & "'"
		if not .eof then 
			LoadValues(prsCoupon)
			FindBycpCouponCode = True
		end if
	end with
	
End Function

'***********************************************************************************************

Public Function LoadBycpCouponCode(strcpCouponCode)

dim sql
dim rs

	On Error Resume Next

	sql = "Select * from sfCoupons where cpCouponCode = '" & strcpCouponCode & "'"
	set rs = server.CreateObject("adodb.recordset")
	set rs = pConnection.Execute(sql)
	
	
	if not (rs.EOF or rs.BOF) then 
		LoadValues(rs)
		LoadBycpCouponCode = True
	end if
	rs.Close
	set rs = Nothing
	
End Function	'LoadBycpCouponCode

Public Function LoadAll

'	On Error Resume Next

	set prsCoupon = server.CreateObject("adodb.recordset")
	with prsCoupon
		.ActiveConnection = pConnection
		.CursorLocation = 2 'adUseClient
		.CursorType = 3 'adOpenStatic
		.LockType = 1 'adLockReadOnly
		.Source = "Select * from sfCoupons"  & mstrsqlWhere
		.Open
		If Not .EOF Then Call LoadValues(prsCoupon)
	end with

	LoadAll = (err.number = 0)
	
End Function	'LoadAll

'***********************************************************************************************

Public Function Find(strcpCouponCode)

'On Error Resume Next

    With prsCoupon
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(strcpCouponCode) <> 0 Then
                .Find "cpCouponCode = '" & strcpCouponCode & "'"
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues(prsCoupon)
        End If
    End With

End Function    'Find

'***********************************************************************************************

Public Function Delete(strcpCouponCode)

dim sql

On Error Resume Next

	if len(cpCouponCode)=0 or cpCouponCode="0" then Exit Function

	sql = "Delete from sfCoupons where cpCouponCode = '" & strcpCouponCode & "'"
	pConnection.Execute sql,,128
	
	If Err.number = 0 Then
		pstrMessage = strcpCouponCode & " was successfully deleted."
	Else
		pstrMessage = "Error: " & Err.number & " - " & Err.description & "<br />"
	End If
	
	Delete = (Err.number=0)
	
End Function

'***********************************************************************************************

Public Function Update

dim strErrorMessage
dim sql
dim rs
dim blnAdd

	Call LoadFromRequest
	strErrorMessage = ValidateValues
	If len(strErrorMessage) = 0 then
		if len(pstrcpCouponCode) = 0 then pstrcpCouponCode = ""
		set rs = server.CreateObject("adodb.recordset")
		sql = "Select * from sfCoupons where cpCouponCode = '" & pstrcpCouponCode & "'"
		rs.open sql, pConnection, 1,	3
		
		blnAdd = False
		If rs.eof then 
			rs.addnew
			blnAdd = True
		End IF

    	rs.Fields("cpCouponCode").Value = pstrcpCouponCode
    	rs.Fields("cpValue").Value = pdblcpValue
    	rs.Fields("cpType").Value = pstrcpType
    	rs.Fields("cpMin").Value = pdblcpMin
    	rs.Fields("cpDescription").Value = pstrcpDescription
    	rs.Fields("cpActivate").Value = pbytcpActivate
    	rs.Fields("cpNeverExpire").Value = pbytcpNeverExpire
    	If Len(pdtcpExpirationDate) = 0 Then
 	    	rs.Fields("cpExpirationDate").Value = Null
	   	Else
	    	rs.Fields("cpExpirationDate").Value = pdtcpExpirationDate
		End If

		On Error Resume Next

    	rs.Update
   	
    	if Err.number = -2147217887 then
    		if Err.description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." then
				pstrMessage = "<H4>The Promo Code you entered is already in use.<br />Please enter a different code.</H4><br />"	
				pblnError = True
    		end if
    	elseif Err.number = -2147217873 then 
			pstrMessage = "<H4>The Promo Code you entered is already in use.<br />Please enter a different code.</H4><br />"	
			pblnError = True
    	elseif Err.number <> 0 then 
			Response.Write "Error: " & Err.number & " - " & Err.description & "<br />"
'    		Err.Raise Err.number,"modPromo.Update",Err.description
    	end if
    	
    	rs.close
    	set rs = nothing
    	
   		if Err.number=0 then
    		if  blnAdd then
    			pstrMessage = pstrcpDescription & " was successfully added."
    		else
    			pstrMessage = "The changes to " & pstrcpDescription & " were successfully saved."
    		end if
    	end if
	End If
	
	Update = strErrorMessage
	
End Function 'Update

'***********************************************************************************************

Private Sub LoadValues(rs)

    pdtcpExpirationDate = rs("cpExpirationDate")
    pstrcpCouponCode = trim(rs("cpCouponCode"))
    pdblcpValue = rs("cpValue")
    pstrcpType = rs("cpType")
    pdblcpMin = rs("cpMin")
    pstrcpDescription = trim(rs("cpDescription"))
	pbytcpActivate = rs("cpActivate")
	pbytcpNeverExpire = rs("cpNeverExpire")
    
End Sub

Private Sub LoadFromRequest

	with Request.Form
		pdtcpExpirationDate = trim(.Item("cpExpirationDate"))
		pstrcpCouponCode = trim(.Item("cpCouponCode"))
		pdblcpValue = trim(.Item("cpValue"))
		
		If .Item("cpType") = "on" Then
			pstrcpType = "Percent"
		Else
			pstrcpType = "Amount"
		End If
		
		pdblcpMin = trim(.Item("cpMin"))
		pstrcpDescription = trim(.Item("cpDescription"))
		pbytcpActivate =  (.Item("cpActivate") = "on") * -1
		pbytcpNeverExpire =  (.Item("cpNeverExpire") = "on") * -1
	end with
	
End Sub

Private Function ValidateValues

dim strError

	if len(pstrcpDescription)=0 then
		strError = strError & "Please enter a title for the coupon;"
	end if
	if len(pstrcpCouponCode)=0 then
		strError = strError & "Please enter a coupon code;"
	end if
	if not isDate(pdtcpExpirationDate) and len(pdtcpExpirationDate)<>0 then
		strError = strError & "Please enter a valid date for the End Date;"
	end if
	if not isNumeric(pdblcpValue) and len(pdblcpValue)<>0 then
		strError = strError & "Please enter a number for the cpValue amount;"
	elseif len(pdblcpValue)=0 then
		strError = strError & "Please enter a cpValue amount;"
	end if
	if not isNumeric(pdblcpMin) and len(pdblcpMin)<>0 then
		strError = strError & "Please enter a number for the minimum subTotal;"
	elseif len(pdblcpMin)=0 then
		strError = strError & "Please enter a minimum subTotal;"
	end if

	pstrMessage = strError
	pblnError = (len(strError)<>0)
	ValidateValues = strError

End Function	'ValidateValues

'***********************************************************************************************

Public Sub OutputSummary()

Dim i
Dim aSortHeader(3,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr

	With Response

    .Write "<table class='tbl' id='tblSummary' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke'>"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='40%'>"
    .Write "<COLGROUP align='center' width='30%'>"
    .Write "  <tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort by Code in descending order"
		aSortHeader(2,0) = "Sort by Title in descending order"
		aSortHeader(3,0) = "Sort by Coupon in descending order"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort by Code in ascending order"
		aSortHeader(2,0) = "Sort by Title in ascending order"
		aSortHeader(3,0) = "Sort by Coupon  Sites in ascending order"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "&nbsp;&nbsp;Code"
	aSortHeader(2,1) = "Title"
	aSortHeader(3,1) = "Coupon"

	if len(mstrOrderBy) > 0 Then
		pstrOrderBy = mstrOrderBy
	Else
		pstrOrderBy = "1"
	End If
	
	.Write "<TH>&nbsp;</TH>"
	for i = 1 to 3
		If cInt(pstrOrderBy) = i Then
			If (pstrSortOrder = "ASC") Then
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'DESC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='Images/up.gif' border=0 align=bottom></TH>" & vbCrLf
			Else
				.Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'ASC');" & chr(34) & _
								" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
								" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & _
								"&nbsp;&nbsp;&nbsp;<img src='Images/down.gif' border=0 align=bottom></TH>" & vbCrLf
			End If
		Else
		    .Write "  <TH style='cursor:hand;' onclick=" & chr(34) & "SortColumn(" & i & ",'" & pstrSortOrder & "');" & chr(34) & _
							" onMouseOver='HighlightColor(this); return DisplayTitle(this);' onMouseOut='deHighlightColor(this); ClearTitle();'" & _
							" title='" & aSortHeader(i,0) & "'>" & aSortHeader(i,1) & "</TH>" & vbCrLf
		End If
	next 'i

    .Write "  </tr>"
	.Write "<tr><td colspan=4>"
    .Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
	.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' " _
				 & ">"
    .Write "<COLGROUP align='center' width='10%'>"
    .Write "<COLGROUP align='left' width='20%'>"
    .Write "<COLGROUP align='center' width='40%'>"
    .Write "<COLGROUP align='center' width='30%'>"
				  				  
	If prsCoupon.RecordCount > 0 Then
        prsCoupon.MoveFirst
        For i = 1 To prsCoupon.RecordCount
			pstrAbbr = Trim(prsCoupon("cpCouponCode"))
 			pstrTitle = "Click to edit " & pstrAbbr & "."
			pstrURL = "sfCouponAdmin.asp?Action=View&CouponID=" & pstrAbbr

			if pstrAbbr = pstrcpCouponCode then
        		.Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
				.Write "    <TD>&nbsp;</TD>" & vbcrlf
				.Write "    <TD>" & prsCoupon("cpCouponCode") & "</TD>" & vbcrlf
			else
				.Write "<TR title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "ViewDetail('" & pstrAbbr & "')" & chr(34) & ">"
				.Write "    <TD>&nbsp;</TD>" & vbcrlf
				.Write "    <TD><A href='sfCouponAdmin.asp?Action=View&PromotionID=" & pstrAbbr & "'>" & prsCoupon("cpCouponCode") & "</A></TD>" & vbcrlf
        	end if
			.Write "    <TD>&nbsp;" & prsCoupon("cpDescription") & "</TD>" & vbcrlf
			.Write "    <TD>&nbsp;" & prsCoupon("cpValue") & "</TD>" & vbcrlf

            Response.Write "</TR>" & vbCrLf
            prsCoupon.MoveNext
        Next
    Else
        Response.Write "<TR><TD align=center><h3>There are no sfCoupons</h3></TD></TR>"
    End If
    .Write "</td></tr></TABLE></div>"
    .Write "</TABLE>"
	End With
	
End Sub      'OutputSummary

End Class	'clsPromotion

'***********************************************************************************************

Sub DebugPrint(strField,strValue)
	Response.Write "<H3>" & strField & ": " & strValue & "</H3><br />"
End Sub

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

Sub LoadFilter

dim pstrOrderBy

	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")
	blnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")

'Build Filter

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'CouponCode
			pstrOrderBy = "cpCouponCode"
		Case "2"	'Title
			pstrOrderBy = "cpDescription"
		Case "3"	'Promotion
			pstrOrderBy = "cpValue"
	End Select	

	If len(pstrOrderBy) > 0 then
		mstrsqlWhere = mstrsqlWhere & " Order By " & pstrOrderBy & " " & mstrSortOrder
	Else
		mstrsqlWhere = ""
	End If
	
End Sub    'LoadFilter

'***********************************************************************************************
'***********************************************************************************************

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsPromo
Dim mstrOrderBy, mstrSortOrder, blnShowSummary
Dim mstrCouponCode
Dim mstrsqlWhere

mstrPageTitle = "Coupon Administration"

	Call LoadFilter
    mstrCouponCode = Request.QueryString("cpCouponCode")
    If len(mstrCouponCode) = 0 Then mstrCouponCode = Request.Form("cpCouponCode")

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
	set mclsPromo = new clsCoupon
    mclsPromo.Connection = cnn
    
    Select Case mAction
        Case "New", "Update"
            mclsPromo.Update
            If mclsPromo.LoadAll Then mclsPromo.Find mstrCouponCode
        Case "Delete"
            mclsPromo.Delete mstrCouponCode
            mclsPromo.LoadAll
        Case "View"
            If mclsPromo.LoadAll Then mclsPromo.Find mstrCouponCode
        Case "Activate", "Deactivate"
            mclsPromo.Activate mstrCouponCode, mAction= "Activate"
            If mclsPromo.LoadAll Then mclsPromo.Find mstrCouponCode
        Case Else
            mclsPromo.LoadAll
    End Select
    
    With mclsPromo

		Call WriteHeader("body_onload();",True)
%>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/calendar.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var blnIsDirty;
var strModifiedOn;
blnIsDirty = false;

function body_onload()
{
	theForm = document.frmData;
	theKeyField = theForm.cpCouponCode;
<% If blnShowSummary then Response.Write "DisplaySummary();" & vbcrlf %>
}

function btnNew_onclick() 
{

	theForm.Action.value = "New";
	theForm.cpExpirationDate.value = "";
	theForm.cpCouponCode.value = "";
	theForm.cpValue.value = "0";
	theForm.cpMin.value = "0";
	theForm.cpDescription.value = "";
	theForm.cpType.checked = false;
	theForm.cpNeverExpire.checked = false;

	theForm.btnUpdate.value = "Add Coupon";
	theForm.btnDelete.disabled = true;
	
	theForm.cpDescription.focus();

}

function btnDelete_onclick() 
{
var blnConfirm;

	blnConfirm = confirm("Are you sure you wish to delete " + theForm.cpDescription.value + "?");
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

function btnReset_onclick() 
{
	theForm.Action.value = "Update";
	theForm.btnUpdate.value = "Save Changes";
	theForm.btnDelete.disabled = false;
}

function ValidInput() 
{
  if (isEmpty(theForm.cpDescription,"Please enter a title for the Coupon.")) {return(false);}
  if (isEmpty(theForm.cpCouponCode,"Please enter a Coupon code.")) {return(false);}
  if (!isNumeric(theForm.cpValue,false,"Please enter a value for the cpValue.")) {return(false);}
  if (!isNumeric(theForm.cpMin,false,"Please enter a value for the minimum subTotal.")) {return(false);}

return(true);
}

function SortColumn(strColumn,strSortOrder)
{
	theForm.Action.value = "";
	theForm.OrderBy.value = strColumn;
	theForm.SortOrder.value = strSortOrder;
	theForm.submit();
	return false;
}

function ViewDetail(theValue)
{
	theKeyField.value = theValue;
	theForm.Action.value = "View";
	theForm.submit();
	return false;
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

<FORM action='sfCouponAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=Action name=Action value='Update'>
<input type=hidden id=blnShowSummary name=blnShowSummary value=''>
<input type=hidden id=OrderBy name=OrderBy value='<%= mstrOrderBy %>'>
<input type=hidden id=SortOrder name=SortOrder value='<%= mstrSortOrder %>'>

<%= .OutputSummary %>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  
  <TR>
    <TD>&nbsp;<LABEL id=lblcpCouponCode for=cpCouponCode>Coupon Code</LABEL></TD>
    <TD>&nbsp;<INPUT id=cpCouponCode name=cpCouponCode Value="<%= mclsPromo.cpCouponCode %>"></TD>
    <TD>&nbsp;<FONT color=Red>*Required</FONT></TD></TR>
  <TR>
    <TD>&nbsp;<LABEL id=lblcpDescription for=cpDescription>Description</LABEL></TD>
    <TD>&nbsp;<INPUT id=cpDescription name=cpDescription style="HEIGHT: 22px; WIDTH: 496px" Value="<%= mclsPromo.cpDescription %>" size=50 maxlength=100></TD>
    <TD>&nbsp;<FONT color=Red>*Required</FONT></TD></TR>
  <TR>
    <TD>&nbsp;<LABEL id=lblcpExpirationDate for=cpExpirationDate>End Date</LABEL></TD>
    <TD>&nbsp;<INPUT id=cpExpirationDate name=cpExpirationDate Value="<%= mclsPromo.cpExpirationDate %>">&nbsp;
		<A HREF="javascript:doNothing()" title="Select end date"
		onClick="setDateField(theForm.cpExpirationDate); top.newWin = window.open('calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
		<IMG SRC="Images/calendar.gif" BORDER=0></A>&nbsp;(mm/dd/yyyy)
    </TD>
    <TD>&nbsp;</TD></TR>
  <TR>
    <TD>&nbsp;<LABEL id=lblcpValue for=cpValue>Amount</LABEL></TD>
    <TD>&nbsp;<INPUT id=cpValue name=cpValue Value="<%= mclsPromo.cpValue %>" onblur="return isNumeric(theForm.cpValue,true,'Please enter a value for the cpValue.');">&nbsp;&nbsp;&nbsp;
        <INPUT id=cpType name=cpType type=checkbox <% if (mclsPromo.cpType="Percent") then Response.Write "Checked" %>>
        <LABEL id=lblcpType for=cpType>Check if a Percentage off</LABEL>
    </TD>
    <TD>&nbsp;<FONT color=Red>*Required</FONT></TD></TR>
  <TR>
    <TD>&nbsp;<LABEL id=lblcpMin for=cpMin>Minimum Order</LABEL></TD>
    <TD>&nbsp;<INPUT id=cpMin name=cpMin Value="<%= mclsPromo.cpMin %>" onblur="return isNumeric(theForm.cpMin,true,'Please enter a value for the minimum subTotal.');"></TD>
    <TD>&nbsp;<FONT color=Red>*Required</FONT></TD></TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>&nbsp;<INPUT id=cpNeverExpire name=cpNeverExpire type=checkbox <% if (mclsPromo.cpNeverExpire=1) then Response.Write "Checked" %>>
    <LABEL id=lblcpNeverExpire for=cpNeverExpire>Never Expires</LABEL></TD>
    <TD>&nbsp;</TD></TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>&nbsp;<INPUT id=cpActivate name=cpActivate type=checkbox <% if (mclsPromo.cpActivate=1) then Response.Write "Checked" %>>
    <LABEL id=lblcpActivate for=cpActivate>Active</LABEL></TD>
    <TD>&nbsp;</TD></TR>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
		<INPUT class='butn' id=btnNew name=btnNew type=button value=New onclick="return btnNew_onclick()">&nbsp;
		<INPUT class='butn' name=btnReset type=reset value=Reset onclick="return btnReset_onclick()">&nbsp;&nbsp; 
		<INPUT class='butn' name=btnDelete type=button value=Delete onclick="return btnDelete_onclick()"> 
		<INPUT class='butn' name=btnUpdate type=submit value="Save Changes"> 
	</TD>
    <TD>&nbsp;</TD>
  </TR>
</TABLE>
</FORM>
</CENTER>

</BODY>
</HTML>
<% 

	End With
	
	set mclsPromo = Nothing
	set cnn = Nothing

Response.Flush
%>
