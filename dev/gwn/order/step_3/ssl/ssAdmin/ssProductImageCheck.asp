<%Option Explicit
'********************************************************************************
'*   Product Image Checker For StoreFront 6.0
'*   Release Version:	1.00.001
'*   Release Date:		September 16, 2004
'*   Revision Date:		September 16, 2004
'*
'*   Release Notes:                                                             *
'*   -- See Product Documentation                                               *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2004 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True
Server.ScriptTimeout = 900

'/////////////////////////////////////////////////
'/

Const cstrSmallImageField = "prodImageSmallPath"
Const cstrLargeImageField = "prodImageLargePath"
Const cstrCustomFields = ""	'must be comma separated with a leading comma: ex. ", Field1, Field 2"
Const cstrDefaultSmallImagePattern = "images/{prodID}.jpg"	'Note the field should be {Field1}
Const cstrDefaultLargeImagePattern = "images/{prodID}.jpg"
Const cbytDefaultOnlyShowErrors = 1	'can be -1, 0 or 1
Const cblnUpdateImageEvenIfNonExistent = True	'True	'False

'/
'/////////////////////////////////////////////////
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'***********************************************************************************************

Function Load(byRef prsProducts)

dim pstrSQL
dim p_strWhere
dim i
dim sql

	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient
'        .CursorType = 3 'adOpenStatic
'        .CursorType = 1 'adOpenKeySet	- Have to use KeySet for SQL Server

		pstrSQL = "SELECT prodID, prodName, " & cstrSmallImageField & ", " & cstrLargeImageField & cstrCustomFields _
				& " FROM sfProducts" _
				& " Order BY prodID, prodName"

		'debugprint "pstrSQL",pstrSQL
		'Response.Flush	  
		  
		On Error Resume Next
		If Err.number <> 0 Then Err.Clear
		
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		
		If Err.number <> 0 Then
			Response.Write "<h4>Error " & Err.number & ": " & Err.Description & "</h4>"
			Response.Write "<h4>SQL: " & pstrSQL & "</h4>"
			Response.Flush
			Err.Clear
			Load = False
			Exit Function
		End If
		On Error Goto 0
		
	End With

    Load = (Not prsProducts.EOF)

End Function    'Load

'******************************************************************************************************************************************************************

Function warningColor(byVal status)
	Select Case status
		Case enNoImage		: warningColor = "yellow"
		Case enInvalidPath	: warningColor = "red"
		Case enValidPath	: warningColor = "green"
		Case Else			: warningColor = "pink"
	End Select
End Function	'warningColor
  
'******************************************************************************************************************************************************************

Function LoadImageArray(byRef aryProduct)

Dim prsProducts

    If Load(prsProducts) Then
		With prsProducts
			mlngNumProducts = .RecordCount
			ReDim aryProduct(mlngNumProducts)
			For i = 1 To mlngNumProducts
				aryProduct(i) = Array(Trim(.Fields("prodID").Value & ""), _
									Trim(.Fields("prodName").Value & ""), _
									Trim(.Fields(cstrSmallImageField).Value & ""), _
									enNoImage, _
									replaceField(prsProducts, mstrSmallImagePattern), _
									enNoImage, _
									Trim(.Fields(cstrLargeImageField).Value & ""), _
									enNoImage, _
									replaceField(prsProducts, mstrLargeImagePattern), _
									enNoImage, _
									enNoImage)
				.MoveNext
			Next 'i
		End With	'prsProducts
	End If	'Load(prsProducts)
	Call ReleaseObject(prsProducts)

	cbytSmallPos = 2
	cbytLargePos = 6
	cbytSummaryPos = 10

End Function	'LoadImageArray
  
'******************************************************************************************************************************************************************

Sub ValidateImagePaths(byRef aryProduct)

Const pstrPathToRemove = "../"
Dim pobjFSO
Dim pstrPathToCheck

	Set pobjFSO = server.CreateObject("Scripting.FileSystemObject")
	For i = 1 To mlngNumProducts
	
		pbytPosToCheck = cbytSmallPos
		If Len(aryProduct(i)(pbytPosToCheck)) > 0 Then
			If Left(aryProduct(i)(pbytPosToCheck), Len(pstrPathToRemove)) = pstrPathToRemove Then
				pstrPathToCheck = cstrBaseFilePath & Replace(aryProduct(i)(pbytPosToCheck), pstrPathToRemove, "", 1, 1)
			Else
				pstrPathToCheck = cstrBaseFilePath & aryProduct(i)(pbytPosToCheck)
			End If
			If pobjFSO.FileExists(pstrPathToCheck) Then
				aryProduct(i)(pbytPosToCheck+1) = enValidPath
			Else
				aryProduct(i)(pbytPosToCheck+1) = enInvalidPath
			End If
		End If
		
		pbytPosToCheck = pbytPosToCheck + 2
		If Len(aryProduct(i)(pbytPosToCheck)) > 0 Then
			If Left(aryProduct(i)(pbytPosToCheck), Len(pstrPathToRemove)) = pstrPathToRemove Then
				pstrPathToCheck = cstrBaseFilePath & Replace(aryProduct(i)(pbytPosToCheck), pstrPathToRemove, "", 1, 1)
			Else
				pstrPathToCheck = cstrBaseFilePath & aryProduct(i)(pbytPosToCheck)
			End If
			If pobjFSO.FileExists(pstrPathToCheck) Then
				aryProduct(i)(pbytPosToCheck+1) = enValidPath
			Else
				aryProduct(i)(pbytPosToCheck+1) = enInvalidPath
			End If
		End If
	
		pbytPosToCheck = cbytLargePos
		If Len(aryProduct(i)(pbytPosToCheck)) > 0 Then
			If Left(aryProduct(i)(pbytPosToCheck), Len(pstrPathToRemove)) = pstrPathToRemove Then
				pstrPathToCheck = cstrBaseFilePath & Replace(aryProduct(i)(pbytPosToCheck), pstrPathToRemove, "", 1, 1)
			Else
				pstrPathToCheck = cstrBaseFilePath & aryProduct(i)(pbytPosToCheck)
			End If
			If pobjFSO.FileExists(pstrPathToCheck) Then
				aryProduct(i)(pbytPosToCheck+1) = enValidPath
			Else
				aryProduct(i)(pbytPosToCheck+1) = enInvalidPath
			End If
		End If
		
		pbytPosToCheck = pbytPosToCheck + 2
		If Len(aryProduct(i)(pbytPosToCheck)) > 0 Then
			If Left(aryProduct(i)(pbytPosToCheck), Len(pstrPathToRemove)) = pstrPathToRemove Then
				pstrPathToCheck = cstrBaseFilePath & Replace(aryProduct(i)(pbytPosToCheck), pstrPathToRemove, "", 1, 1)
			Else
				pstrPathToCheck = cstrBaseFilePath & aryProduct(i)(pbytPosToCheck)
			End If
			If pobjFSO.FileExists(pstrPathToCheck) Then
				aryProduct(i)(pbytPosToCheck+1) = enValidPath
			Else
				aryProduct(i)(pbytPosToCheck+1) = enInvalidPath
			End If
		End If
		
		If CBool(aryProduct(i)(cbytSmallPos+1) = enInvalidPath) OR CBool(aryProduct(i)(cbytLargePos+1) = enInvalidPath) Then
			aryProduct(i)(cbytSummaryPos) = enInvalidPath
		ElseIf CBool(aryProduct(i)(cbytSmallPos+1) = enNoImage) OR CBool(aryProduct(i)(cbytLargePos+1) = enNoImage) Then
			aryProduct(i)(cbytSummaryPos) = enNoImage
		Else
			aryProduct(i)(cbytSummaryPos) = enValidPath
		End If
	Next 'i
	Set pobjFSO = Nothing

End Sub	'ValidateImagePaths
  
'******************************************************************************************************************************************************************

Function updateImagePaths(byRef aryProduct)

Dim pstrSQL
Dim pbytPosToCheck
Dim pstrFieldName
Dim pblnEmptyIfInvalid
Dim pstrResult
Dim pstrTempResult

	pblnEmptyIfInvalid = True	'True	'False
	
	On Error Resume Next
	
	For i = 1 To mlngNumProducts
		pstrSQL = ""
		pstrTempResult = ""
		
		pstrFieldName = cstrSmallImageField
		pbytPosToCheck = cbytSmallPos
		Select Case mbytSmallImageUpdate
			Case 1: 'Invalid paths only
				If CBool(aryProduct(i)(pbytPosToCheck+1) = enInvalidPath) Then
					If CBool(aryProduct(i)(pbytPosToCheck+3) = enValidPath Or cblnUpdateImageEvenIfNonExistent) Then
						pstrSQL = makeSQLUpdate(pstrFieldName, aryProduct(i)(pbytPosToCheck+2), False, enDatatype_string)
					ElseIf CBool(aryProduct(i)(pbytPosToCheck+3) = enInvalidPath) And pblnEmptyIfInvalid Then
						pstrSQL = makeSQLUpdate(pstrFieldName, "", True, enDatatype_string)
					Else
						pstrTempResult = pstrTempResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & pstrFieldName & " image <em>" & aryProduct(i)(pbytPosToCheck+2) & "</em> does not exist - NO UPDATE ATTEMPTED</font></li>"
					End If
				End If
			Case 2: 'Invalid or undefined paths
				If (CBool(aryProduct(i)(pbytPosToCheck+1) = enInvalidPath) Or CBool(aryProduct(i)(pbytPosToCheck+1) = enNoImage)) Then
					If CBool(aryProduct(i)(pbytPosToCheck+3) = enValidPath Or cblnUpdateImageEvenIfNonExistent) Then
						pstrSQL = makeSQLUpdate(pstrFieldName, aryProduct(i)(pbytPosToCheck+2), False, enDatatype_string)
					ElseIf CBool(aryProduct(i)(pbytPosToCheck+3) = enInvalidPath) And pblnEmptyIfInvalid Then
						pstrSQL = makeSQLUpdate(pstrFieldName, "", True, enDatatype_string)
					Else
						pstrTempResult = pstrTempResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & pstrFieldName & " image <em>" & aryProduct(i)(pbytPosToCheck+2) & "</em> does not exist - NO UPDATE ATTEMPTED</font></li>"
					End If
				End If
			Case 3: 'Update all
				If CBool(aryProduct(i)(pbytPosToCheck+3) = enValidPath Or cblnUpdateImageEvenIfNonExistent) Then
					pstrSQL = makeSQLUpdate(pstrFieldName, aryProduct(i)(pbytPosToCheck+2), False, enDatatype_string)
				ElseIf CBool(aryProduct(i)(pbytPosToCheck+3) = enInvalidPath) And pblnEmptyIfInvalid Then
					pstrSQL = makeSQLUpdate(pstrFieldName, "", True, enDatatype_string)
				Else
					pstrTempResult = pstrTempResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & pstrFieldName & " image <em>" & aryProduct(i)(pbytPosToCheck+2) & "</em> does not exist - NO UPDATE ATTEMPTED</font></li>"
				End If
		End Select
		
		pstrFieldName = cstrLargeImageField
		pbytPosToCheck = cbytLargePos
		Select Case mbytLargeImageUpdate
			Case 1: 'Invalid paths only
				If CBool(aryProduct(i)(pbytPosToCheck+1) = enInvalidPath) Then
					If CBool(aryProduct(i)(pbytPosToCheck+3) = enValidPath Or cblnUpdateImageEvenIfNonExistent) Then
						pstrSQL = makeSQLUpdate(pstrFieldName, aryProduct(i)(pbytPosToCheck+2), False, enDatatype_string)
					ElseIf CBool(aryProduct(i)(pbytPosToCheck+3) = enInvalidPath) And pblnEmptyIfInvalid Then
						pstrSQL = makeSQLUpdate(pstrFieldName, "", True, enDatatype_string)
					Else
						pstrTempResult = pstrTempResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & pstrFieldName & " image <em>" & aryProduct(i)(pbytPosToCheck+2) & "</em> does not exist - NO UPDATE ATTEMPTED</font></li>"
					End If
				End If
			Case 2: 'Invalid or undefined paths
				If (CBool(aryProduct(i)(pbytPosToCheck+1) = enInvalidPath) Or CBool(aryProduct(i)(pbytPosToCheck+1) = enNoImage)) Then
					If CBool(aryProduct(i)(pbytPosToCheck+3) = enValidPath Or cblnUpdateImageEvenIfNonExistent) Then
						pstrSQL = makeSQLUpdate(pstrFieldName, aryProduct(i)(pbytPosToCheck+2), False, enDatatype_string)
					ElseIf CBool(aryProduct(i)(pbytPosToCheck+3) = enInvalidPath) And pblnEmptyIfInvalid Then
						pstrSQL = makeSQLUpdate(pstrFieldName, "", True, enDatatype_string)
					Else
						pstrTempResult = pstrTempResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & pstrFieldName & " image <em>" & aryProduct(i)(pbytPosToCheck+2) & "</em> does not exist - NO UPDATE ATTEMPTED</font></li>"
					End If
				End If
			Case 3: 'Update all
				If CBool(aryProduct(i)(pbytPosToCheck+3) = enValidPath Or cblnUpdateImageEvenIfNonExistent) Then
					pstrSQL = makeSQLUpdate(pstrFieldName, aryProduct(i)(pbytPosToCheck+2), False, enDatatype_string)
				ElseIf CBool(aryProduct(i)(pbytPosToCheck+3) = enInvalidPath) And pblnEmptyIfInvalid Then
					pstrSQL = makeSQLUpdate(pstrFieldName, "", True, enDatatype_string)
				Else
					pstrTempResult = pstrTempResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & pstrFieldName & " image <em>" & aryProduct(i)(pbytPosToCheck+2) & "</em> does not exist - NO UPDATE ATTEMPTED</font></li>"
				End If
		End Select
		
		If Len(pstrSQL) > 0 Then
			If Left(pstrSQL, 2) = ", " Then pstrSQL = Replace(pstrSQL, ", ", "", 1, 1)
			pstrSQL = "Update sfProducts Set " & pstrSQL & " Where prodID=" & wrapSQLValue(aryProduct(i)(0), False, enDatatype_string)
			cnn.Execute pstrSQL,,128
			'Response.Write i & ": " & pstrSQL & "<br />"
			
			If Err.number = 0 Then
				pstrResult = pstrResult & "<li>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & " updated."
			Else
				pstrResult = pstrResult & "<li><font color=red>" & aryProduct(i)(0) & ": " & aryProduct(i)(1) & " - Error " & Err.number & ": " & Err.Description & "</font></li>"
				Err.Clear
			End If
		End If
		
		pstrResult = pstrResult & pstrTempResult

    Next 'i
    
    If Len(pstrResult) > 0 Then pstrResult = "<li>" & pstrResult & "</li>"
    
    mstrMessage = pstrResult
    'Response.Write pstrResult

End Function	'updateImagePaths
  
'******************************************************************************************************************************************************************

mstrPageTitle = "Product Image Check"

Const enNoImage = 0
Const enInvalidPath = -1
Const enValidPath = 1

Dim cstrBaseFilePath
Dim pstrFilePath
Dim i
Dim mbytOnlyShowErrors
Dim mstrAction

Dim cbytSmallPos
Dim mbytSmallImageUpdate
Dim mstrSmallImagePattern

Dim cbytLargePos
Dim mbytLargeImageUpdate
Dim mstrLargeImagePattern

Dim cbytSummaryPos

Dim mlngNumProducts
Dim pstrFileToCheck
Dim maryProduct
Dim pbytPosToCheck
'Dim mstrMessage
  
	cstrBaseFilePath = PhysPath
	'debugprint "cstrBaseFilePath", cstrBaseFilePath
	
	mbytOnlyShowErrors = LoadRequestValue("OnlyShowErrors")
	If Len(mbytOnlyShowErrors) = 0 Then
		mbytOnlyShowErrors = cbytDefaultOnlyShowErrors
	ElseIf isNumeric(mbytOnlyShowErrors) Then
		mbytOnlyShowErrors = CLng(mbytOnlyShowErrors)
	Else
		mbytOnlyShowErrors = cbytDefaultOnlyShowErrors
	End If
  
	mbytSmallImageUpdate = LoadRequestValue("SmallImageUpdate")
	If Len(mbytSmallImageUpdate) = 0 Then
		mbytSmallImageUpdate = 0
	ElseIf isNumeric(mbytSmallImageUpdate) Then
		mbytSmallImageUpdate = CLng(mbytSmallImageUpdate)
	Else
		mbytSmallImageUpdate = 0
	End If
	mstrSmallImagePattern = LoadRequestValue("SmallImagePattern")
	
	mbytLargeImageUpdate = LoadRequestValue("LargeImageUpdate")
	If Len(mbytLargeImageUpdate) = 0 Then
		mbytLargeImageUpdate = 0
	ElseIf isNumeric(mbytLargeImageUpdate) Then
		mbytLargeImageUpdate = CLng(mbytLargeImageUpdate)
	Else
		mbytLargeImageUpdate = 0
	End If
	mstrLargeImagePattern = LoadRequestValue("LargeImagePattern")
	
	mstrAction = LoadRequestValue("btnSubmit")
	'debugprint "mstrAction", mstrAction
	
	If Len(mstrAction) = 0 Then
		If Len(mstrSmallImagePattern) = 0 Then mstrSmallImagePattern = cstrDefaultSmallImagePattern
		If Len(mstrLargeImagePattern) = 0 Then mstrLargeImagePattern = cstrDefaultLargeImagePattern
	End If

	Call LoadImageArray(maryProduct)
	Call ValidateImagePaths(maryProduct)
	If mstrAction = "Update" Then
		Call updateImagePaths(maryProduct)
		Call LoadImageArray(maryProduct)
		Call ValidateImagePaths(maryProduct)
	End If
	
	Call ReleaseObject(cnn)
		
	Call WriteHeader("",True)
%>
<CENTER>
<form action="ssProductImageCheck.asp" name="frmData" id="frmData" method=post>


<table border=0 cellPadding=5 cellSpacing=1 width="100%" ID="tblMain">
  <tr>
    <th><div class="pagetitle "><%= mstrPageTitle %></div></th>
</table>

<% If Len(mstrMessage) > 0 Then Response.Write "<fieldset><legend>Result</legend><div align=left>" & mstrMessage & "</div></fieldset>" %>

<table class="tbl" width="100%" cellpadding="2" cellspacing="0" border="1" id="tblSummary">
  <tr>
    <td colspan="4">
      <table>
        <tr>
          <td>
            <fieldset>
              <legend>Display images with</legend>
			&nbsp;&nbsp;<input type="radio" name="OnlyShowErrors" id="OnlyShowErrors0" value="<%= enInvalidPath %>" <%= isChecked(mbytOnlyShowErrors=enInvalidPath) %>>&nbsp;<label for="OnlyShowErrors0">Invalid paths only</label><br />
			&nbsp;&nbsp;<input type="radio" name="OnlyShowErrors" id="OnlyShowErrors1" value="<%= enNoImage %>" <%= isChecked(mbytOnlyShowErrors=enNoImage) %>>&nbsp;<label for="OnlyShowErrors1">Invalid or undefined paths</label><br />
			&nbsp;&nbsp;<input type="radio" name="OnlyShowErrors" id="OnlyShowErrors2" value="<%= enValidPath %>" <%= isChecked(mbytOnlyShowErrors=enValidPath) %>>&nbsp;<label for="OnlyShowErrors2">Show all</label>
            </fieldset>
          </td>
        </tr>
        <tr>
          <td>
            <fieldset>
              <legend>Update small images with</legend>
			&nbsp;&nbsp;<input type="radio" name="SmallImageUpdate" id="SmallImageUpdate0" value="0" <%= isChecked(mbytSmallImageUpdate=0) %>>&nbsp;<label for="SmallImageUpdate0">Do not update</label><br />
			&nbsp;&nbsp;<input type="radio" name="SmallImageUpdate" id="SmallImageUpdate1" value="1" <%= isChecked(mbytSmallImageUpdate=1) %>>&nbsp;<label for="SmallImageUpdate1">Invalid paths only</label><br />
			&nbsp;&nbsp;<input type="radio" name="SmallImageUpdate" id="SmallImageUpdate2" value="2" <%= isChecked(mbytSmallImageUpdate=2) %>>&nbsp;<label for="SmallImageUpdate2">Invalid or undefined paths</label><br />
			&nbsp;&nbsp;<input type="radio" name="SmallImageUpdate" id="SmallImageUpdate3" value="3" <%= isChecked(mbytSmallImageUpdate=3) %>>&nbsp;<label for="SmallImageUpdate3">Update all</label><br />
			<input type=text name="SmallImagePattern" id="SmallImagePattern" value="<%= mstrSmallImagePattern %>" size="20">
            </fieldset>
          </td>
        </tr>
        <tr>
          <td>
            <fieldset>
              <legend>Update large images with</legend>
			&nbsp;&nbsp;<input type="radio" name="LargeImageUpdate" id="LargeImageUpdate0" value="0" <%= isChecked(mbytLargeImageUpdate=0) %>>&nbsp;<label for="LargeImageUpdate0">Do not update</label><br />
			&nbsp;&nbsp;<input type="radio" name="LargeImageUpdate" id="LargeImageUpdate1" value="1" <%= isChecked(mbytLargeImageUpdate=1) %>>&nbsp;<label for="LargeImageUpdate1">Invalid paths only</label><br />
			&nbsp;&nbsp;<input type="radio" name="LargeImageUpdate" id="LargeImageUpdate2" value="2" <%= isChecked(mbytLargeImageUpdate=2) %>>&nbsp;<label for="LargeImageUpdate2">Invalid or undefined paths</label><br />
			&nbsp;&nbsp;<input type="radio" name="LargeImageUpdate" id="LargeImageUpdate3" value="3" <%= isChecked(mbytLargeImageUpdate=3) %>>&nbsp;<label for="LargeImageUpdate3">Update all</label><br />
			<input type=text name="LargeImagePattern" id="Text1" value="<%= mstrLargeImagePattern %>" size="20">
            </fieldset>
          </td>
        </tr>
        <tr>
          <td><input type=submit name="btnSubmit" id="btnSubmit" value="Verify">&nbsp;<input type=submit name="btnSubmit" id="btnSubmit1" value="Update"></td>
        </tr>
        <tr>
          <td>
            <fieldset>
              <legend>Color Codes</legend>
              <table cellpadding=2 cellspacing=0 border=1>
                <tr><td bgcolor=<%= warningColor(enInvalidPath) %>>&nbsp;&nbsp;&nbsp;</td><td>Invalid Path - Image defined in database but not present</td></tr>
                <tr><td bgcolor=<%= warningColor(enNoImage) %>>&nbsp;&nbsp;&nbsp;</td><td>Undefined Path - No image defined in database</td></tr>
                <tr><td bgcolor=<%= warningColor(enValidPath) %>>&nbsp;&nbsp;&nbsp;</td><td>Valid Path - Image defined in database and present</td></tr>
              </table>
            </fieldset>
          </td>
        </tr>
      </table>
    
    </td>
  </tr>
  <tr class="tblhdr">
    <th colspan="1" align="left">Code</th>
    <th colspan="1" align="left">Product</th>
    <th colspan="1" align="left">Small Image</th>
    <th colspan="1" align="left">Large Image</th>
  </tr>

  <% 
	For i = 1 To mlngNumProducts
		If maryProduct(i)(cbytSummaryPos) <= mbytOnlyShowErrors Then
 %>
  <tr>
    <td colspan="1" align="left"><a href="sfProductAdmin.asp?Action=ViewProduct&ViewID=<%= maryProduct(i)(0) %>"><%= maryProduct(i)(0) %></a>&nbsp;</td>
    <td colspan="1" align="left"><a href="sfProductAdmin.asp?Action=ViewProduct&ViewID=<%= maryProduct(i)(0) %>"><%= maryProduct(i)(1) %>&nbsp;</td>
	<td colspan="1" align="left" bgcolor="<%= warningColor(maryProduct(i)(cbytSmallPos+1)) %>"><%= maryProduct(i)(cbytSmallPos) %></td>
	<td colspan="1" align="left" bgcolor="<%= warningColor(maryProduct(i)(cbytLargePos+1)) %>"><%= maryProduct(i)(cbytLargePos) %></td>
  </tr>
  <%
		End If
    Next 'i
  %>
</TABLE>

</FORM>
<!--#include file="adminFooter.asp"-->
</CENTER>
</BODY>
</HTML>
<% If Response.Buffer Then Response.Flush %>