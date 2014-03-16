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

Class clssfText
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private pRS
Private pblnError
'database variables

Private plngID
Private pstrtxtCategoryIsActive
Private pstrtxtCategoryP
Private pstrtxtCategoryS
Private pstrtxtDateIsActive
Private pstrtxtDes
Private pstrtxtDesign
Private pstrtxtManufacturerP
Private pstrtxtManufacturerS
Private pstrtxtMFGIsActive
Private pstrtxtPrice
Private pstrtxtPriceIsActive
Private pstrtxtProdId
Private pstrtxtQnty
Private pstrtxtSaleIsActive
Private pstrtxtSaleP
Private pstrtxtVendorIsActive
Private pstrtxtVendorP
Private pstrtxtVendorS
Private pstrtxtYSave

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


Public Property Get ID()
    ID = plngID
End Property

Public Property Get txtCategoryIsActive()
    txtCategoryIsActive = pstrtxtCategoryIsActive
End Property

Public Property Get txtCategoryP()
    txtCategoryP = pstrtxtCategoryP
End Property

Public Property Get txtCategoryS()
    txtCategoryS = pstrtxtCategoryS
End Property

Public Property Get txtDateIsActive()
    txtDateIsActive = pstrtxtDateIsActive
End Property

Public Property Get txtDes()
    txtDes = pstrtxtDes
End Property

Public Property Get txtDesign()
    txtDesign = pstrtxtDesign
End Property

Public Property Get txtManufacturerP()
    txtManufacturerP = pstrtxtManufacturerP
End Property

Public Property Get txtManufacturerS()
    txtManufacturerS = pstrtxtManufacturerS
End Property

Public Property Get txtMFGIsActive()
    txtMFGIsActive = pstrtxtMFGIsActive
End Property

Public Property Get txtPrice()
    txtPrice = pstrtxtPrice
End Property

Public Property Get txtPriceIsActive()
    txtPriceIsActive = pstrtxtPriceIsActive
End Property

Public Property Get txtProdId()
    txtProdId = pstrtxtProdId
End Property

Public Property Get txtQnty()
    txtQnty = pstrtxtQnty
End Property

Public Property Get txtSaleIsActive()
    txtSaleIsActive = pstrtxtSaleIsActive
End Property

Public Property Get txtSaleP()
    txtSaleP = pstrtxtSaleP
End Property

Public Property Get txtVendorIsActive()
    txtVendorIsActive = pstrtxtVendorIsActive
End Property

Public Property Get txtVendorP()
    txtVendorP = pstrtxtVendorP
End Property

Public Property Get txtVendorS()
    txtVendorS = pstrtxtVendorS
End Property

Public Property Get txtYSave()
    txtYSave = pstrtxtYSave
End Property

'***********************************************************************************************

Private Sub LoadValues(rs)

    plngID = trim(rs("ID"))
    pstrtxtCategoryIsActive = (trim(rs("txtCategoryIsActive")) = "1")
    pstrtxtCategoryP = trim(rs("txtCategoryP"))
    pstrtxtCategoryS = trim(rs("txtCategoryS"))
    pstrtxtDateIsActive = (trim(rs("txtDateIsActive")) = "1")
    pstrtxtDes = trim(rs("txtDes"))
    pstrtxtDesign = trim(rs("txtDesign"))
    pstrtxtManufacturerP = trim(rs("txtManufacturerP"))
    pstrtxtManufacturerS = trim(rs("txtManufacturerS"))
    pstrtxtMFGIsActive = (trim(rs("txtMFGIsActive")) = "1")
    pstrtxtPrice = trim(rs("txtPrice"))
    pstrtxtPriceIsActive = (trim(rs("txtPriceIsActive")) = "1")
    pstrtxtProdId = trim(rs("txtProdId"))
    pstrtxtQnty = trim(rs("txtQnty"))
    pstrtxtSaleIsActive = (trim(rs("txtSaleIsActive")) = "1")
    pstrtxtSaleP = trim(rs("txtSaleP"))
    pstrtxtVendorIsActive = (trim(rs("txtVendorIsActive")) = "1")
    pstrtxtVendorP = trim(rs("txtVendorP"))
    pstrtxtVendorS = trim(rs("txtVendorS"))
    pstrtxtYSave = trim(rs("txtYSave"))

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
        plngID = Trim(.Item("ID"))
        pstrtxtCategoryIsActive = (uCase(.Item("txtCategoryIsActive")) = "ON")
        pstrtxtCategoryP = Trim(.Item("txtCategoryP"))
        pstrtxtCategoryS = Trim(.Item("txtCategoryS"))
        pstrtxtDateIsActive = (uCase(.Item("txtDateIsActive")) = "ON")
        pstrtxtDes = Trim(.Item("txtDes"))
        pstrtxtDesign = Trim(.Item("txtDesign"))
        pstrtxtManufacturerP = Trim(.Item("txtManufacturerP"))
        pstrtxtManufacturerS = Trim(.Item("txtManufacturerS"))
        pstrtxtMFGIsActive = (uCase(.Item("txtMFGIsActive")) = "ON")
        pstrtxtPrice = Trim(.Item("txtPrice"))
        pstrtxtPriceIsActive = (uCase(.Item("txtPriceIsActive")) = "ON")
        pstrtxtProdId = Trim(.Item("txtProdId"))
        pstrtxtQnty = Trim(.Item("txtQnty"))
        pstrtxtSaleIsActive = (uCase(.Item("txtSaleIsActive")) = "ON")
        pstrtxtSaleP = Trim(.Item("txtSaleP"))
        pstrtxtVendorIsActive = (uCase(.Item("txtVendorIsActive")) = "ON")
        pstrtxtVendorP = Trim(.Item("txtVendorP"))
        pstrtxtVendorS = Trim(.Item("txtVendorS"))
        pstrtxtYSave = Trim(.Item("txtYSave"))
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function LoadAll()

'On Error Resume Next

    Set pRS = GetRS("Select * from sfText")
    If Not (pRS.EOF Or pRS.BOF) Then
        Call LoadValues(pRS)
        LoadAll = True
    End If

End Function    'LoadAll

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
        If Len(plngID) = 0 Then plngID = 0

        sql = "Select * from sfText where ID = " & plngID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("txtCategoryIsActive") = (pstrtxtCategoryIsActive * -1)
        rs("txtCategoryP") = pstrtxtCategoryP
        rs("txtCategoryS") = pstrtxtCategoryS
        rs("txtDateIsActive") = (pstrtxtDateIsActive * -1)
        rs("txtDes") = pstrtxtDes
        rs("txtDesign") = pstrtxtDesign
        rs("txtManufacturerP") = pstrtxtManufacturerP
        rs("txtManufacturerS") = pstrtxtManufacturerS
        rs("txtMFGIsActive") = (pstrtxtMFGIsActive * -1)
        rs("txtPrice") = pstrtxtPrice
        rs("txtPriceIsActive") = (pstrtxtPriceIsActive * -1)
        rs("txtProdId") = pstrtxtProdId
        rs("txtQnty") = pstrtxtQnty
        rs("txtSaleIsActive") = (pstrtxtSaleIsActive * -1)
        rs("txtSaleP") = pstrtxtSaleP
        rs("txtVendorIsActive") = (pstrtxtVendorIsActive * -1)
        rs("txtVendorP") = pstrtxtVendorP
        rs("txtVendorS") = pstrtxtVendorS
        rs("txtYSave") = pstrtxtYSave

        rs.Update

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<BR>Please enter a different data.</H4><BR>"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<BR>"
        End If
        
        plngID = rs("ID")
        rs.Close
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = "The record was successfully added."
            Else
                pstrMessage = "Your settings were successfully saved."
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

'***********************************************************************************************

Public Sub WriteDesignToFile()

'On Error Resume Next

Dim fso, MyFile
Dim p_strFile

	p_strFile = mstrBasePath & "SFLib/incText.asp"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	On Error Resume Next
	
	Set MyFile = fso.CreateTextFile(p_strFile, True)
	
	If Err.number = 0 Then
		With MyFile
			.WriteLine "<" & "%"
			.WriteLine "'******************************************************************"
			.WriteLine "' "
			.WriteLine "' Created With Sandshot Software's WebStore Manager for Storefront "
			.WriteLine "' "
			.WriteLine "' Design constants used for StoreFront Search Engine "
			.WriteLine "' "
			.WriteLine "'*******************************************************************"
			.WriteLine ""
			.WriteLine "	Const C_CategoryNameS     = " & chr(34) & pstrtxtCategoryS & chr(34)
			.WriteLine "	Const C_CategoryNameP     = " & chr(34) & pstrtxtCategoryP & chr(34)
			.WriteLine "	Const C_ManufacturerNameS = " & chr(34) & pstrtxtManufacturerS & chr(34)
			.WriteLine "	Const C_ManufacturerNameP = " & chr(34) & pstrtxtManufacturerP & chr(34)
			.WriteLine "	Const C_VendorNameS       = " & chr(34) & pstrtxtVendorS & chr(34)
			.WriteLine "	Const C_VendorNameP       = " & chr(34) & pstrtxtVendorP & chr(34)
			.WriteLine "	Const C_ProductID         = " & chr(34) & pstrtxtProdId & chr(34)
			.WriteLine "	Const C_Description       = " & chr(34) & pstrtxtDes & chr(34)
			.WriteLine "	Const C_Price             = " & chr(34) & pstrtxtPrice & chr(34)
			.WriteLine "	Const C_QUANTITY          = " & chr(34) & pstrtxtQnty & chr(34)
			.WriteLine "	Const C_SPrice            = " & chr(34) & pstrtxtSaleP & chr(34)
			.WriteLine "	Const C_YSave             = " & chr(34) & pstrtxtYSave & chr(34)
			.WriteLine "	Const C_DesignType        = " & chr(34) & pstrtxtDesign & chr(34)
			.WriteLine "	Const C_CategoryIsActive  = " & (pstrtxtCategoryIsActive * -1)
			.WriteLine "	Const C_VendorIsActive    = " & (pstrtxtVendorIsActive * -1)
			.WriteLine "	Const C_MFGIsActive       = " & (pstrtxtMFGIsActive * -1)
			.WriteLine "	Const C_AddedIsActive     = " & (pstrtxtDateIsActive * -1)
			.WriteLine "	Const C_PriceIsActive     = " & (pstrtxtPriceIsActive * -1)
			.WriteLine "	Const C_SaleIsActive      = " & (pstrtxtSaleIsActive * -1)
			.WriteLine "%" & ">"
			
			.Close
		End With
	Else
		If err.Description = "Permission denied" Then
			Response.Write "<h3><font color=red>Unable to save settings to design file SFLib/incText.asp</font></h3>"
			Response.Write "<p>This error occurs when the IUSR_<em>xxxx</em> does not have write permissions. Please contact your host to resolve this setting.</p>"
			Response.Write "<p>The changes you just made will <b>Not</b> show up in the live site until this is fixed. You may manually edit the file SFLib/incText.asp.</p>"
		Else
			Response.Write "<h3><font color=red>Unable to save settings to design file SFLib/incText.asp</font></h3>"
			Response.Write "<p>Error " & err.number & ": " & err.Description & "</p>"
			Response.Write "<p>The changes you just made will <b>Not</b> show up in the live site until this is fixed. You may manually edit the file SFLib/incText.asp.</p>"
		End If

		Err.Clear
	End If

	Set fso = Nothing
	Set MyFile = Nothing

End Sub      'WriteDesignToFile

'***********************************************************************************************

End Class   'clssfText

mstrPageTitle = "Search Engine Design"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclssfText

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Set mclssfText = New clssfText
    
    Select Case mAction
        Case "New", "Update"
            mclssfText.Update
            mclssfText.WriteDesignToFile
        Case Else
            mclssfText.LoadAll
    End Select
    
	Call WriteHeader("",True)
    With mclssfText
%>

<SCRIPT LANGUAGE=javascript>
<!--

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

<FORM action='sfTextAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=ID name=ID value=<%= .ID %>>
<input type=hidden id=Action name=Action value='Update'>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none">
<tr class='tblhdr'>
<th colspan=1 align=left>&nbsp;</th>
<th colspan=1 align=left><span id="spanDetailTitle">Labels for Search and Search Output Pages</span></th>
<th colspan=1 align=left><span id="spanDetailTitle">Advanced Search Options</span></th>
</tr>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtSaleP for=txtSaleP>Sale Price:</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtSaleP name=txtSaleP Value='<%= .txtSaleP %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=txtSaleIsActive name=txtSaleIsActive <% If .txtSaleIsActive then Response.Write "checked" %>>&nbsp;<LABEL id=lbltxtSaleIsActive for=txtSaleIsActive>Permit Searches for Sale Items</LABEL></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtPrice for=txtPrice>Price:</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtPrice name=txtPrice Value='<%= .txtPrice %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=txtPriceIsActive name=txtPriceIsActive <% If .txtPriceIsActive then Response.Write "checked" %>>&nbsp;<LABEL id=lbltxtPriceIsActive for=txtPriceIsActive>Permit Searches by Price</LABEL></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtYSave for=txtYSave>"You Save":</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtYSave name=txtYSave Value='<%= .txtYSave %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=txtMFGIsActive name=txtMFGIsActive <% If .txtMFGIsActive then Response.Write "checked" %>>&nbsp;<LABEL id=lbltxtMFGIsActive for=txtMFGIsActive>Permit Searches by Manufacturer</LABEL></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtDes for=txtDes>Description:</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtDes name=txtDes Value='<%= .txtDes %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=txtVendorIsActive name=txtVendorIsActive <% If .txtVendorIsActive then Response.Write "checked" %>>&nbsp;<LABEL id=lbltxtVendorIsActive for=txtVendorIsActive>Permit Searches by Vendor</LABEL></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtProdId for=txtProdId>Product ID:</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtProdId name=txtProdId Value='<%= .txtProdId %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=txtCategoryIsActive name=txtCategoryIsActive <% If .txtCategoryIsActive then Response.Write "checked" %>>&nbsp;<LABEL id=lbltxtCategoryIsActive for=txtCategoryIsActive>Permit Searches by Category</LABEL></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtQnty for=txtQnty>Quantity:</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtQnty name=txtQnty Value='<%= .txtQnty %>' maxlength=50 size=50></TD>
        <TD>&nbsp;<INPUT type=checkbox id=txtDateIsActive name=txtDateIsActive <% If .txtDateIsActive then Response.Write "checked" %>>&nbsp;<LABEL id=lbltxtDateIsActive for=txtDateIsActive>Permit Searches by Date Added</LABEL></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtCategoryS for=txtCategoryS>Category:</LABEL></TD>
        <TD colspan=2>&nbsp;<INPUT id=txtCategoryS name=txtCategoryS Value='<%= .txtCategoryS %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtCategoryP for=txtCategoryP>Category (plural):</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtCategoryP name=txtCategoryP Value='<%= .txtCategoryP %>' maxlength=50 size=50></TD>
      </TR>
	 <TR><TD Colspan=3><HR></TD></TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtManufacturerS for=txtManufacturerS>Manufacturer:</LABEL></TD>
        <TD colspan=2>&nbsp;<INPUT id=txtManufacturerS name=txtManufacturerS Value='<%= .txtManufacturerS %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtManufacturerP for=txtManufacturerP>Manufacturer (plural):</LABEL></TD>
        <TD colspan=2>&nbsp;<INPUT id=txtManufacturerP name=txtManufacturerP Value='<%= .txtManufacturerP %>' maxlength=50 size=50></TD>
      </TR>
	 <TR><TD Colspan=3><HR></TD></TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtVendorS for=txtVendorS>Vendor:</LABEL></TD>
        <TD colspan=2>&nbsp;<INPUT id=txtVendorS name=txtVendorS Value='<%= .txtVendorS %>' maxlength=50 size=50></TD>
      </TR>
      <TR>
        <TD class="Label">&nbsp;<LABEL id=lbltxtVendorP for=txtVendorP>Vendor (plural):</LABEL></TD>
        <TD>&nbsp;<INPUT id=txtVendorP name=txtVendorP Value='<%= .txtVendorP %>' maxlength=50 size=50></TD>
      </TR>
	 <TR><TD Colspan=3><HR></TD></TR>
      <TR>
        <TD class="Label">&nbsp;Image Layout:</LABEL></TD>
        <TD colspan=2>
          &nbsp;<input type=radio name=txtDesign id=txtDesign0 Value=1 <%= isChecked(CStr(.txtDesign) = "1") %>>&nbsp;<LABEL for=txtDesign0>Left</LABEL><br>
          &nbsp;<input type=radio name=txtDesign id=txtDesign1 Value=2 <%= isChecked(CStr(.txtDesign) = "2") %>>&nbsp;<LABEL for=txtDesign1>Right</LABEL><br>
          &nbsp;<input type=radio name=txtDesign id=txtDesign2 Value=3 <%= isChecked(CStr(.txtDesign) = "3") %>>&nbsp;<LABEL for=txtDesign2>Alternating</LABEL>
        </TD>
      </TR>
  <TR>
    <TD colspan=3 align=center>
        <INPUT class="butn" id=btnUpdate name=btnUpdate type=submit value='Save Changes'>
    </TD>
  </TR>
</TABLE>
</FORM>

</CENTER>
</BODY>
</HTML>
<%
    End With
    Set mclssfText = Nothing
    Set cnn = Nothing
    Response.Flush
%>
