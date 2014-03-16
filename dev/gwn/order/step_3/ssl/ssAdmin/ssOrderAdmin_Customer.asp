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

Class clssfCustomers
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private pblnBillingAddress

Private pstrcustAddr1
Private pstrcustAddr2
Private pstrcustCity
Private pstrcustCountry
Private pstrcustEmail
Private pstrcustFAX
Private plngcustID
Private pstrcustLastName
Private pstrcustFirstName
Private pstrcustMiddleInitial
Private pstrcustCompany
Private pstrcustPhone
Private pstrcustState
Private pstrcustZip

Private pblncustIsSubscribed
Private plngcustTimesAccessed
Private pdtcustLastAccess

Private plngPricingLevelID


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


Public Property Get custAddr1()
    custAddr1 = pstrcustAddr1
End Property

Public Property Get custAddr2()
    custAddr2 = pstrcustAddr2
End Property

Public Property Get custCity()
    custCity = pstrcustCity
End Property

Public Property Get custCountry()
    custCountry = pstrcustCountry
End Property

Public Property Get custEmail()
    custEmail = pstrcustEmail
End Property

Public Property Get custFAX()
    custFAX = pstrcustFAX
End Property

Public Property Get custID()
    custID = plngcustID
End Property

Public Property Get custFirstName()
    custFirstName = pstrcustFirstName
End Property

Public Property Get custMiddleInitial()
    custMiddleInitial = pstrcustMiddleInitial
End Property

Public Property Get custLastName()
    custLastName = pstrcustLastName
End Property

Public Property Get custCompany()
    custCompany = pstrcustCompany
End Property

Public Property Get custPhone()
    custPhone = pstrcustPhone
End Property

Public Property Get custState()
    custState = pstrcustState
End Property

Public Property Get custZip()
    custZip = pstrcustZip
End Property

Public Property Get custIsSubscribed()
   custIsSubscribed  = pblncustIsSubscribed
End Property

Public Property Get custTimesAccessed()
    custTimesAccessed = plngcustTimesAccessed
End Property

Public Property Get custLastAccess()
    custLastAccess = pdtcustLastAccess
End Property

Public Property Get PricingLevelID()
    PricingLevelID = plngPricingLevelID
End Property

Public Property Let BillingAddress(blnBillingAddress)
	pblnBillingAddress = blnBillingAddress
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

	If pblnBillingAddress Then
		plngcustID = trim(rs.Fields("custID").Value)
		pstrcustFirstName = trim(rs.Fields("custFirstName").Value)
		pstrcustMiddleInitial = trim(rs.Fields("custMiddleInitial").Value)
		pstrcustLastName = trim(rs.Fields("custLastName").Value)
		pstrcustCompany = trim(rs.Fields("custCompany").Value)
		pstrcustAddr1 = trim(rs.Fields("custAddr1").Value)
		pstrcustAddr2 = trim(rs.Fields("custAddr2").Value)
		pstrcustCity = trim(rs.Fields("custCity").Value)
		pstrcustState = trim(rs.Fields("custState").Value)
		pstrcustZip = trim(rs.Fields("custZip").Value)
		pstrcustCountry = trim(rs.Fields("custCountry").Value)
		pstrcustEmail = trim(rs.Fields("custEmail").Value)
		pstrcustPhone = trim(rs.Fields("custPhone").Value)
		pstrcustFAX = trim(rs.Fields("custFAX").Value)

		pblncustIsSubscribed = trim(rs.Fields("custIsSubscribed").Value)
		plngcustTimesAccessed = trim(rs.Fields("custTimesAccessed").Value)
		pdtcustLastAccess = trim(rs.Fields("custLastAccess").Value)

		If cblnAddon_PricingLevelMgr Then plngPricingLevelID = trim(rs.Fields("PricingLevelID").Value)
	Else
		plngcustID = trim(rs.Fields("cshpaddrID").Value)
		pstrcustFirstName = trim(rs.Fields("cshpaddrShipFirstName").Value)
		pstrcustMiddleInitial = trim(rs.Fields("cshpaddrShipMiddleInitial").Value)
		pstrcustLastName = trim(rs.Fields("cshpaddrShipLastName").Value)
		pstrcustCompany = trim(rs.Fields("cshpaddrShipCompany").Value)
		pstrcustAddr1 = trim(rs.Fields("cshpaddrShipAddr1").Value)
		pstrcustAddr2 = trim(rs.Fields("cshpaddrShipAddr2").Value)
		pstrcustCity = trim(rs.Fields("cshpaddrShipCity").Value)
		pstrcustState = trim(rs.Fields("cshpaddrShipState").Value)
		pstrcustZip = trim(rs.Fields("cshpaddrShipZip").Value)
		pstrcustCountry = trim(rs.Fields("cshpaddrShipCountry").Value)
		pstrcustEmail = trim(rs.Fields("cshpaddrShipEmail").Value)
		pstrcustPhone = trim(rs.Fields("cshpaddrShipPhone").Value)
		pstrcustFAX = trim(rs.Fields("cshpaddrShipFAX").Value)
	End If

End Sub 'LoadValues


Private Sub LoadFromRequest

    With Request.Form
        pstrcustAddr1 = Trim(.Item("custAddr1"))
        pstrcustAddr2 = Trim(.Item("custAddr2"))
        pstrcustCity = Trim(.Item("custCity"))
        pstrcustCountry = Trim(.Item("custCountry"))
        pstrcustEmail = Trim(.Item("custEmail"))
        pstrcustFAX = Trim(.Item("custFAX"))
        plngcustID = Trim(.Item("ID"))
        pstrcustFirstName = Trim(.Item("custFirstName"))
        pstrcustMiddleInitial = Trim(.Item("custMiddleInitial"))
        pstrcustLastName = Trim(.Item("custLastName"))
        pstrcustCompany = Trim(.Item("custCompany"))
        pstrcustPhone = Trim(.Item("custPhone"))
        pstrcustState = Trim(.Item("custState"))
        pstrcustZip = Trim(.Item("custZip"))

        pblncustIsSubscribed = (lCase(.Item("custIsSubscribed")) = "on")

		plngPricingLevelID = Trim(.Item("PricingLevelID"))

    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function Load(ByRef lngID)

Dim pstrSQL

'On Error Resume Next

	If pblnBillingAddress Then
		pstrSQL = "Select * from sfCustomers where custID=" & lngID
	Else
		pstrSQL = "Select * from sfCShipAddresses Where cshpaddrID=" & lngID
	End If
    Set pRS = GetRS(pstrSQL)
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        Load = True
    End If

End Function    'Load

'***********************************************************************************************

Public Function Update()

Dim pstrSQL
Dim rs
Dim strErrorMessage
Dim blnAdd

'On Error Resume Next

    pblnError = False
    Call LoadFromRequest

    If Len(plngcustID) = 0 Then
		pstrMessage = "Please provide a customer ID"
        Update = False
        Exit Function
    End If

    strErrorMessage = ValidateValues
    If ValidateValues Then

		If pblnBillingAddress Then
			pstrSQL = "Select * from sfCustomers where custID=" & plngcustID
		Else
			pstrSQL = "Select * from sfCShipAddresses Where cshpaddrID=" & plngcustID
		End If
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open pstrSQL, cnn, 1, 3
        If rs.EOF Then
			pstrMessage = "Could Not Locate the customer record"
            Update = False
            Exit Function
        End If

		If pblnBillingAddress Then
			rs.Fields("custFirstName") = pstrcustFirstName
			rs.Fields("custMiddleInitial") = pstrcustMiddleInitial
			rs.Fields("custLastName") = pstrcustLastName
			rs.Fields("custCompany") = pstrcustCompany
			rs.Fields("custAddr1") = pstrcustAddr1
			rs.Fields("custAddr2") = pstrcustAddr2
			rs.Fields("custCity") = pstrcustCity
			rs.Fields("custState") = pstrcustState
			rs.Fields("custZip") = pstrcustZip
			rs.Fields("custCountry") = pstrcustCountry
			rs.Fields("custFAX") = pstrcustFAX
			rs.Fields("custPhone") = pstrcustPhone
			rs.Fields("custEmail") = pstrcustEmail
	        
			rs.Fields("custIsSubscribed") = Abs(pblncustIsSubscribed * 1)

			If cblnAddon_PricingLevelMgr Then
				If Len(plngPricingLevelID) = 0 Then
					rs.Fields("PricingLevelID") = Null
				Else
					rs.Fields("PricingLevelID") = plngPricingLevelID
				End If
			End If
		Else
			rs.Fields("cshpaddrShipFirstName") = pstrcustFirstName
			rs.Fields("cshpaddrShipMiddleInitial") = pstrcustMiddleInitial
			rs.Fields("cshpaddrShipLastName") = pstrcustLastName
			rs.Fields("cshpaddrShipCompany") = pstrcustCompany
			rs.Fields("cshpaddrShipAddr1") = pstrcustAddr1
			rs.Fields("cshpaddrShipAddr2") = pstrcustAddr2
			rs.Fields("cshpaddrShipCity") = pstrcustCity
			rs.Fields("cshpaddrShipState") = pstrcustState
			rs.Fields("cshpaddrShipZip") = pstrcustZip
			rs.Fields("cshpaddrShipCountry") = pstrcustCountry
			rs.Fields("cshpaddrShipFAX") = pstrcustFAX
			rs.Fields("cshpaddrShipPhone") = pstrcustPhone
			rs.Fields("cshpaddrShipEmail") = pstrcustEmail
		End If

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrcustLastName & " was successfully added."
            Else
                pstrMessage = "The changes to " & pstrcustLastName & " were successfully saved."
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

Function ValidateValues()

Dim strError

    strError = ""

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues
End Class   'clssfCustomers

mstrPageTitle = "Customer Administration"
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'the following line is added for Version 2
Call WriteHeader("",False)

'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfCustomers
Dim mvntID
Dim mblnBillingAddress

    mvntID = LoadRequestValue("ID")
    mAction = LoadRequestValue("Action")
    mblnBillingAddress = LoadRequestValue("BillingAddress")
    
    If Len(mblnBillingAddress) = 0 Then	mblnBillingAddress = CBool(mAction <> "EditShipping")

    Set mclssfCustomers = New clssfCustomers
	mclssfCustomers.BillingAddress = mblnBillingAddress
    
    Select Case mAction
        Case "Update"
            mclssfCustomers.Update
        Case Else
    End Select
    mclssfCustomers.Load mvntID
    
    With mclssfCustomers
%>

<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function ValidInput(theForm)
{

    return(true);
}
//-->
</SCRIPT>

<CENTER>
<%= .OutputMessage %>
<% If mAction = "Update" Then %>
<h4>You will need to refresh the Order Admin Page to view the updated information.</h4>
<h4>Click <a href="" onclick="window.parent.opener.document.location.reload(); window.close(); return false;">here</a> to continue.</h4>
<center><a href="" onclick="window.close(); return false;">Close</a></center>
<% End If %>
<FORM action='ssOrderAdmin_Customer.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=BillingAddress name=BillingAddress value=<%= mblnBillingAddress %>>
<input type=hidden id="ID" name="ID" value=<%= mvntID %>>
<input type=hidden id=Action name=Action value='Update'>

<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <colgroup align=right>
  <colgroup align=left>
  <tr class='tblhdr'>
	<th colspan="3" align=center><span id="spanDetailTitle"><%= .custLastName %> 
	<% If mblnBillingAddress Then %>
	 - Billing Address
	<% Else %>
	 - Shipping Address
	<% End If %>
	</span></th>
  </tr>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustLastName for=custLastName>Name:</LABEL></TD>
        <TD>
          <INPUT id=custLastName name=custLastName Value="<%= .custLastName %>" maxlength=50 size=50>,&nbsp;
          <INPUT id=custFirstName name=custFirstName Value="<%= .custFirstName %>" maxlength=50 size=50>
          <INPUT id=custMiddleInitial name=custMiddleInitial Value="<%= .custMiddleInitial %>" maxlength=1 size=1>
        </TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustCompany for=custLastName>Company:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custCompany name=custCompany Value="<%= .custCompany %>" maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustAddr1 for=custAddr1>Address 1:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custAddr1 name=custAddr1 Value="<%= .custAddr1 %>" maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustAddr2 for=custAddr2>Address 2:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custAddr2 name=custAddr2 Value="<%= .custAddr2 %>" maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustCity for=custCity>City:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custCity name=custCity Value="<%= .custCity %>" maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustState for=custState>State:</LABEL></TD>
        <td>
			<select size="1"  id="custState" name=custState>
			<% Call MakeCombo("Select loclstAbbreviation, loclstName from sfLocalesState Where loclstLocaleIsActive=1","loclstName","loclstAbbreviation",.custState) %>
			</select>
        </td>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustZip for=custZip>ZIP:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custZip name=custZip Value="<%= .custZip %>" maxlength=15 size=15></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustCountry for=custCountry>Country:</LABEL></TD>
        <TD>&nbsp;
			<select size="1"  id=custCountry name=custCountry>
			<% Call MakeCombo("Select loclctryAbbreviation, loclctryName from sfLocalesCountry Where loclctryLocalIsActive=1","loclctryName","loclctryAbbreviation",.custCountry) %>
			</select>
		</TD>        
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustEmail for=custEmail>Email:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custEmail name=custEmail Value="<%= .custEmail %>" maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustPhone for=custPhone>Phone:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custPhone name=custPhone Value="<%= .custPhone %>" maxlength=20 size=20></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustFAX for=custFAX>FAX:</LABEL></TD>
        <TD>&nbsp;<INPUT id=custFAX name=custFAX Value="<%= .custFAX %>" maxlength=20 size=20></TD>
      </TR>
	  <% If mblnBillingAddress Then %>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustTimesAccessed for=custTimesAccessed>Times Accessed:</LABEL></TD>
        <TD>&nbsp;<%= .custTimesAccessed %></TD>
      </TR>
      <TR>
        <TD>&nbsp;<LABEL id=lblcustLastAccess for=custLastAccess>Last Accessed:</LABEL></TD>
        <TD>&nbsp;<%= .custLastAccess %></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;</TD>
        <TD><INPUT type=checkbox id=custIsSubscribed name=custIsSubscribed <% If .custIsSubscribed Then Response.Write "Checked" %>>&nbsp;<LABEL id=lblcustIsSubscribed for=custIsSubscribed>Subscribed to mailing list.</LABEL></TD>
      </TR>
	  <% If cblnAddon_PricingLevelMgr Then	'Pricing Level Manager %>
      <TR>
        <TD>&nbsp;<LABEL id=lblPricingLevelID for=PricingLevelID>PricingLevelID:</LABEL></TD>
        <TD>&nbsp;
			<select size="1"  id=PricingLevelID name=PricingLevelID>
			<%
			Dim i
			Dim pobjRS

			set pobjRS = GetRS("Select PricingLevelID, PricingLevelName from PricingLevels")
				
			If len(.PricingLevelID) = 0 Then
				Response.Write "<option selected value=" & chr(34) & chr(34) & ">None</option>" & vbcrlf
			Else
				Response.Write "<option value=" & chr(34) & chr(34) & ">None</option>" & vbcrlf
			End If

			For i = 1 to pobjRS.recordcount
				if len(.PricingLevelID) > 0 then
					if trim(pobjRS("PricingLevelID")-1) <> trim(.PricingLevelID) then
						Response.Write "<option value=" & chr(34) & (pobjRS.Fields("PricingLevelID").Value-1) & chr(34) & ">" & pobjRS.Fields("PricingLevelName").Value & "</option>" & vbcrlf
					else
						Response.Write "<option selected value=" & chr(34) & (pobjRS.Fields("PricingLevelID").Value-1) & chr(34) & ">" & pobjRS.Fields("PricingLevelName").Value & "</option>" & vbcrlf
					end if
				else
					Response.Write "<option value=" & chr(34) & (pobjRS.Fields("PricingLevelID").Value-1) & chr(34) & ">" & pobjRS.Fields("PricingLevelName").Value & "</option>" & vbcrlf
				end if
				pobjRS.movenext
			Next
				
			pobjRS.Close
			set pobjRS = nothing

			%>
			</select>
		</TD>        
      </TR>
      <% End If	'Pricing Level Manager %>
      <% End If	'mblnBillingAddress %>
  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)'>&nbsp;&nbsp;
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
    Set mclssfCustomers = Nothing
    Set cnn = Nothing
    Response.Flush
%>
