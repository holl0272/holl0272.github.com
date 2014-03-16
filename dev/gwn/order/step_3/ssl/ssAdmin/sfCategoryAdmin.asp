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

Class clsCategory
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsCategory
Private prssubCategory
Private pblnError

'database variables
Private plngcatID
Private pstrcatName
Private pstrcatDescription
Private pstrcatImage
Private pstrcatHasSubCategory
Private pblncatIsActive
Private pstrcatHttp

Private plngsubcatID
Private pstrsubcatName
Private pstrsubcatDescription
Private pstrsubcatImage
Private pblnsubcatIsActive
Private pblnUseSubCategory

'***********************************************************************************************

Private Sub class_Initialize()
    cstrDelimeter  = ";"
    pblnUseSubCategory = False
End Sub

Private Sub class_Terminate()
    On Error Resume Next
    Set prsCategory = Nothing
    If pblnUseSubCategory Then Set prssubCategory = Nothing
End Sub

'***********************************************************************************************

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

'***********************************************************************************************

Public Property Get UseSubCategory()
    UseSubCategory = pblnUseSubCategory
End Property
Public Property Get catID()
    catID = plngcatID
End Property
Public Property Get catName()
    catName = pstrcatName
End Property
Public Property Get catDescription()
    catDescription = pstrcatDescription
End Property
Public Property Get catImage()
    catImage = pstrcatImage
End Property
Public Property Get catHasSubCategory()
    catHasSubCategory = pstrcatHasSubCategory
End Property
Public Property Get catIsActive()
    catIsActive = pblncatIsActive
End Property
Public Property Get catHttp()
    catHttp = pstrcatHttp
End Property

Public Property Get subcatID()
    subcatID = plngsubcatID
End Property
Public Property Get subcatName()
    subcatName = pstrsubcatName
End Property
Public Property Get subcatDescription()
    subcatDescription = pstrsubcatDescription
End Property
Public Property Get subcatImage()
    subcatImage = pstrsubcatImage
End Property
Public Property Get subcatIsActive()
    subcatIsActive = pblnsubcatIsActive
End Property

Public Property Get rssubCategory()
	If pblnUseSubCategory Then Set rssubCategory = prssubCategory
End Property

'***********************************************************************************************

Private Sub ClearValues()

	plngcatID = ""
	pstrcatName = ""
	pstrcatDescription = ""
	pstrcatImage = ""
	pstrcatHasSubCategory = ""
	pblncatIsActive = False
	pstrcatHttp = ""

	plngsubcatID = ""
	pstrsubcatName = ""
	pstrsubcatDescription = ""
	pstrsubcatImage = ""
	pblnsubcatIsActive = False

End Sub 'ClearValues

Private Sub LoadValues

	plngcatID = Trim(prsCategory("catID"))
	pstrcatName = Trim(prsCategory("catName"))
	pstrcatDescription = Trim(prsCategory("catDescription"))
	pstrcatImage = Trim(prsCategory("catImage"))
	
	If isNull(prsCategory("catHasSubCategory")) Then
		pstrcatHasSubCategory = False
	Else
		pstrcatHasSubCategory = cBool(prsCategory("catHasSubCategory"))
	End If
	If isNull(prsCategory("catIsActive")) Then
		pblncatIsActive = False
	Else
		pblncatIsActive = cBool(prsCategory("catIsActive"))
	End If
	
	pstrcatHttp = Trim(prsCategory("catHttpAdd"))

	If pblnUseSubCategory Then
		If (not prssubCategory.EOF) and (len(mlngsubCatID) > 0) Then
			plngsubcatID = Trim(prssubCategory("subcatID"))
			pstrsubcatName = Trim(prssubCategory("subcatName"))
			pstrsubcatDescription = Trim(prssubCategory("subcatDescription"))
			pstrsubcatImage = Trim(prssubCategory("subcatImage"))
			If isNull(prssubCategory("subcatIsActive")) Then
				pblnsubcatIsActive = False
			Else
				pblnsubcatIsActive = cBool(prssubCategory("subcatIsActive"))
			End If
		Else
			plngsubcatID = ""
			pstrsubcatName = ""
			pstrsubcatDescription = ""
			pstrsubcatImage = ""
			pblnsubcatIsActive = False
		End If
	End If

End Sub 'LoadValues

Private Sub LoadFromRequest

    With Request.Form
		plngcatID = Trim(.Item("catID"))
		pstrcatName = Trim(.Item("catName"))
		pstrcatDescription = Trim(.Item("catDescription"))
		pstrcatImage = Trim(.Item("catImage"))
		pstrcatHasSubCategory = Trim(.Item("catHasSubCategory"))
		pblncatIsActive = (uCase(.Item("catIsActive"))) = "ON"
		pstrcatHttp = Trim(.Item("catHttp"))

		plngsubcatID = Trim(.Item("subcatID"))
		pstrsubcatName = Trim(.Item("subcatName"))
		pstrsubcatDescription = Trim(.Item("subcatDescription"))
		pstrsubcatImage = Trim(.Item("subcatImage"))
		pblnsubcatIsActive = uCase(Trim(.Item("subcatIsActive"))) = "ON"
    End With

End Sub 'LoadFromRequest

'***********************************************************************************************

Public Function FindCategory(lngCatID)

'On Error Resume Next

    With prsCategory
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngCatID) <> 0 Then
                .Find "CatID=" & lngCatID
            Else
                .MoveLast
            End If
            If Not .EOF Then LoadValues
        End If
    End With

End Function    'FindCategory

Public Function FindsubCategory(lngsubCatID)

'On Error Resume Next

    With prssubCategory
        If .RecordCount > 0 Then
            .MoveFirst
            If Len(lngsubCatID) <> 0 Then
                .Find "subCatID=" & lngsubCatID
            Else
                .MoveLast
            End If
            If Not .EOF Then Call FindCategory(prssubCategory("subcatCategoryId"))
        End If
    End With

End Function    'FindsubCategory

'***********************************************************************************************

Public Function Load()

'On Error Resume Next

    Set prsCategory = GetRS("Select * from sfCategories " & mstrsqlWhere)
    If pblnUseSubCategory Then Set prssubCategory = GetRS("Select * from sfsubCategories Order By subcatName")
    If not prsCategory.EOF Then LoadValues
    Load = (Not prsCategory.EOF)

End Function    'Load

'***********************************************************************************************

Public Function ActivateCategory(lngID,blnActivate)

Dim sql

On Error Resume Next

	if blnActivate then
		sql = "Update sfCategories Set catIsActive=1 where catID=" & lngID
        pstrMessage = "Category successfully activated."
    else
		sql = "Update sfCategories Set catIsActive=0 where catID=" & lngID
        pstrMessage = "Category successfully deactivated."
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        ActivateCategory = True
    Else
        pstrMessage = Err.Description
        ActivateCategory = False
    End If

End Function    'ActivateCategory

Public Function ActivatesubCategory(lngID,blnActivate)

Dim sql

On Error Resume Next

	if blnActivate then
		sql = "Update sfsubCategories Set subCatIsActive=1 where subcatID=" & lngID
        pstrMessage = "sub-category successfully activated."
    else
		sql = "Update sfsubCategories Set subCatIsActive=0 where subcatID=" & lngID
        pstrMessage = "sub-category successfully deactivated."
	end if
    cnn.Execute sql, , 128
    
    If (Err.Number = 0) Then
        ActivatesubCategory = True
    Else
        pstrMessage = Err.Description
        ActivatesubCategory = False
    End If

End Function    'ActivatesubCategory

'***********************************************************************************************

Public Function DeleteCategory(lngID)

Dim sql

'On Error Resume Next

	If len(lngID) = 0 Then Exit Function
    sql = "Delete from sfsubCategories where subcatCategoryId = " & lngID
    cnn.Execute sql, , 128
    sql = "Delete from sfCategories where CatID = " & lngID
    cnn.Execute sql, , 128
    If (Err.Number = 0) Then
        pstrMessage = "The category was successfully deleted."
        DeleteCategory = True
    Else
        pstrMessage = Err.Description
        DeleteCategory = False
    End If

End Function    'DeleteCategory

Public Function DeletesubCategory(lngID)

Dim sql
Dim pRS
Dim p_lngCatID

'On Error Resume Next

	If len(lngID) = 0 Then Exit Function
	sql = "Select subcatCategoryID from sfsubCategories where subcatId = " & lngID
	Set pRS = GetRS(sql)
	If not pRS.EOF Then
		p_lngCatID = pRS("subcatCategoryID")
		sql = "Delete from sfsubCategories where subcatId = " & lngID
		cnn.Execute sql, , 128

		sql = "Select subCatID from sfsubCategories where subcatCategoryID = " & p_lngCatID
		Set pRS = GetRS(sql)
		
		If pRS.EOF Then
			sql = "Update sfCategories Set catHasSubCategory=0 where CatID =" & p_lngCatID
			cnn.Execute sql, , 128
		End If
		
		If (Err.Number = 0) Then
		    pstrMessage = "The sub-category was successfully deleted."
		    DeletesubCategory = True
		Else
		    pstrMessage = Err.Description
		    DeletesubCategory = False
		End If
	Else
		pstrMessage = "This sub-category does not exist."
		DeletesubCategory = False
	End If
	pRS.Close
	Set pRS = Nothing

End Function    'DeletesubCategory

'***********************************************************************************************

Public Function CopyCategory(lngID)

Dim sql
Dim pRS, prsCopy
Dim pstrCopyCat

Dim p_lngCatID
Dim p_strcatName
Dim p_strcatDescription
Dim p_strcatImage
Dim p_blncatIsActive
Dim p_blnHasSubCategory
Dim p_strcatHttp

'On Error Resume Next

	pstrCopyCat = Request.Form("CopyCat")
	sql = "Select * from sfCategories where catID = " & lngID
	Set pRS = server.CreateObject("adodb.Recordset")
	pRS.open sql, cnn, 1, 3
	If Not pRS.EOF Then
	
		p_strcatName = pRS("catName")
		p_strcatDescription = pRS("catDescription")
		p_strcatImage = pRS("catImage")
		p_blncatIsActive = pRS("catIsActive")
		p_blnHasSubCategory = pRS("catHasSubCategory")
		p_strcatHttp = pRS("catHttpAdd")

		pRS.AddNew

		pRS("catName") = pstrCopyCat
		pRS("catDescription") = p_strcatDescription
		pRS("catImage") = p_strcatImage
		pRS("catIsActive") = p_blncatIsActive
		pRS("catHasSubCategory") = p_blnHasSubCategory
		pRS("catHttpAdd") = p_strcatHttp

		pRS.Update
		p_lngCatID = pRS("catID")

		If cBool(p_blnHasSubCategory) Then
			sql = "Select * from sfsubCategories where subcatCategoryID = " & lngID
			Set pRS = GetRS(sql)
			
			If not pRS.EOF Then
				sql = "Select * from sfsubCategories where subcatCategoryID = " & p_lngCatID
				Set prsCopy = server.CreateObject("adodb.Recordset")
				prsCopy.open sql, cnn, 1, 3
				do while not pRS.EOF
					prsCopy.AddNew
					prsCopy("subcatCategoryId") = p_lngcatID
					prsCopy("subcatName") = pRS("subcatName")
					prsCopy("subcatDescription") = pRS("subcatDescription")
					If len(pRS("subcatImage"))>0 Then prsCopy("subcatImage") = pRS("subcatImage")
					prsCopy("subcatIsActive") = pRS("subcatIsActive")
			
					pRS.MoveNext
				loop
				prsCopy.Update
			End If
			prsCopy.Close
			set prsCopy = Nothing
		End If
	End If
	
	pRS.Close
	set pRS = Nothing
	
	mlngCatID = p_lngCatID

		If (Err.Number = 0) Then
		    pstrMessage = pstrCopyCat & " was successfully created."
		    CopyCategory = True
		Else
		    pstrMessage = Err.Description
		    CopyCategory = False
		End If

End Function    'CopyCategory

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
        If Len(plngcatID) = 0 Then plngcatID = 0

        sql = "Select * from sfCategories where catID = " & plngcatID
        Set rs = server.CreateObject("adodb.Recordset")
        rs.open sql, cnn, 1, 3
        If rs.EOF Then
            rs.AddNew
            blnAdd = True
        Else
            blnAdd = False
        End If

        rs("catName") = pstrcatName
        rs("catDescription") = pstrcatDescription
        rs("catImage") = pstrcatImage
        rs("catIsActive") = pblncatIsActive * -1
        rs("catHasSubCategory") = (len(pstrsubcatName) > 0) * -1
        rs("catHttpAdd") = pstrcatHttp
        rs.Update
        plngcatID = rs("catID")
		mlngCatID = plngCatID
        rs.Close

		If (len(pstrsubcatName) > 0) Then
	        If Len(plngsubCatID) = 0 Then plngsubCatID = 0
			sql = "Select * from sfsubCategories where subCatID = " & plngsubCatID
			rs.open sql, cnn, 1, 3
			
			If rs.EOF Then rs.AddNew

			rs("subcatCategoryId") = plngcatID
			rs("subcatName") = pstrsubcatName
			rs("subcatDescription") = pstrsubcatDescription
			If len(pstrsubcatImage)>0 Then rs("subcatImage") = pstrsubcatImage
			rs("subcatIsActive") = pblnsubcatIsActive * -1
			rs.Update
			rs.Close
		End If

        If Err.Number = -2147217887 Then
            If Err.Description = "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again." Then
                pstrMessage = "<H4>The data you entered is already in use.<br />Please enter a different data.</H4><br />"
                pblnError = True
            End If
        ElseIf Err.Number <> 0 Then
            Response.Write "Error: " & Err.Number & " - " & Err.Description & "<br />"
        End If
        
        Set rs = Nothing
        
        If Err.Number = 0 Then
            If blnAdd Then
                pstrMessage = pstrcatName & " was successfully added."
            Else
                pstrMessage = "The changes to " & pstrcatName & " were successfully saved."
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

Public Sub OutputSubCatValues()

Dim i

	If pblnUseSubCategory Then
		With prssubCategory
			if .RecordCount > 0 then .MoveFirst
			If not .EOF Then .filter = "subcatCategoryId=" & prsCategory("CatID")

			Response.Write "var arysubCatID = new Array();" & vbcrlf
			Response.Write "var arysubcatName = new Array();" & vbcrlf
			Response.Write "var arysubcatDescription = new Array();" & vbcrlf
			Response.Write "var arysubcatImage = new Array();" & vbcrlf
			Response.Write "var arysubcatIsActive = new Array();" & vbcrlf
			Response.Write vbcrlf
			For i = 1 to .RecordCount
				Response.Write "arysubCatID[" & i & "] = " & chr(34) & prssubCategory("subcatID") & chr(34) & ";" & vbcrlf
				Response.Write "arysubcatName[" & i & "] = " & chr(34) & prssubCategory("subcatName") & chr(34) & ";" & vbcrlf
				Response.Write "arysubcatDescription[" & i & "] = " & chr(34) & prssubCategory("subcatDescription") & chr(34) & ";" & vbcrlf
				Response.Write "arysubcatImage[" & i & "] = " & chr(34) & prssubCategory("subcatImage") & chr(34) & ";" & vbcrlf
				Response.Write "arysubcatIsActive[" & i & "] = " & lCase(cBool(prssubCategory("subcatIsActive"))) & ";" & vbcrlf
		
  				.MoveNext
			Next
			.filter = ""
		End With
	End If

End Sub

'***********************************************************************************************

Public Sub OutputSummary()

'On Error Resume Next

Dim i
Dim aSortHeader(3,1)
Dim pstrOrderBy, pstrSortOrder
Dim pstrTitle, pstrURL, pstrAbbr
Dim pstrSelect
Dim pblnActive

	With Response

		.Write "<table class='tbl' width='95%' cellpadding='0' cellspacing='0' border='1' rules='none' bgcolor='whitesmoke' id='tblSummary'>"
		.Write "<colgroup align='left' width='5%'>"
		.Write "<colgroup align='left' width='10%'"
		.Write "<colgroup align='left' width='85%'>"
		.Write "<colgroup align='center'>"
		.Write "	<tr class='tblhdr'>"

	If (len(mstrSortOrder) = 0) or (mstrSortOrder = "ASC") Then
		aSortHeader(1,0) = "Sort Categories in descending order"
		aSortHeader(2,0) = "Sort Descriptions in descending order"
		aSortHeader(3,0) = "Sort Active Categories first"
		pstrSortOrder = "ASC"
	Else
		aSortHeader(1,0) = "Sort Categories in ascending order"
		aSortHeader(2,0) = "Sort Descriptions in ascending order"
		aSortHeader(3,0) = "Sort Inactive Categories first"
		pstrSortOrder = "DESC"
	End If
	aSortHeader(1,1) = "Category"
	aSortHeader(2,1) = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Description"
	aSortHeader(3,1) = "Active&nbsp;&nbsp;"

	.Write "<TH>&nbsp;</TH>"
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

		.Write "	</tr>"

		.Write "<tr><td colspan=4>"
		.Write "<div name='divSummary' style='height:100; overflow:scroll;'>"
		.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='1' rules='none' " _
				 & "bgcolor='whitesmoke' style='cursor:hand;' name='tblSummary'" _
				 & ">"
		.Write "<colgroup align='left' width='5%'>"
		.Write "<colgroup align='left' width='25%'>"
		.Write "<colgroup align='left' width='52%'>"
		.Write "<colgroup align='left' width='13%'>"
    If prsCategory.RecordCount > 0 Then
        prsCategory.MoveFirst
        For i = 1 To prsCategory.RecordCount
			pstrURL = "sfCategoryAdmin.asp?Action=View&CatID=" & prsCategory("CatID")
			pstrTitle = "Click to view " & prsCategory("catName")
			pstrSelect = "title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();'"
			
			If (isNull(prsCategory("CatIsActive")) or len(prsCategory("CatIsActive"))=0) Then
				pblnActive = false
			Else
				pblnActive = cBool(prsCategory("CatIsActive"))
			End If
					
            If Trim(prsCategory("CatID")) = plngCatID Then
                .Write "<TR class='Selected' onmouseover='doMouseOverRow(this)' onmouseout='doMouseOutRow(this)'>"
            Else
				if pblnActive then
					.Write " <TR class='Active' " & pstrSelect & ">"
				else
					.Write " <TR class='Inactive' " & pstrSelect & ">"
        		end if
            End If
            
            If pblnUseSubCategory Then
				prssubCategory.filter = "subcatCategoryID=" & prsCategory("CatID")
				if prssubCategory.RecordCount > 0 then
					prssubCategory.MoveFirst
					If (Trim(prsCategory("CatID")) = plngCatID) and (len(mlngsubCatID) <> 0) Then
						.Write " <TD>&nbsp;&nbsp;&nbsp;&nbsp;<span onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='View sub-Categories' onclick='return ExpandCategory(this," & prsCategory("CatID") & ");'>-</span></TD>"
					Else
						.Write " <TD>&nbsp;&nbsp;&nbsp;&nbsp;<span onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='View sub-Categories' onclick='return ExpandCategory(this," & prsCategory("CatID") & ");'>+</span></TD>"
					End If
				else
					.Write " <TD>&nbsp;</TD>"
       			end if
			else
				.Write " <TD>&nbsp;</TD>"
       		end if
            If Trim(prsCategory("CatID")) = plngCatID Then
        		.Write "<TD>" & prsCategory("catName") & "&nbsp;</TD>"
            Else
        		.Write "<TD><a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & prsCategory("catName") & "&nbsp;</a></TD>"
            End If
       		.Write "<TD onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">" & prsCategory("catDescription") & "&nbsp;</TD>"

			if pblnActive then
        		.Write "<TD><a href='sfCategoryAdmin.asp?Action=Deactivate&CatID=" & prsCategory("CatID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & prsCategory("catName") & ".'>Active</a></TD></TR>" & vbCrLf
			else
        		.Write "<TD><a href='sfCategoryAdmin.asp?Action=Activate&CatID=" & prsCategory("CatID") & _
										"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & prsCategory("catName") & ".'>Inactive</a></TD></TR>" & vbCrLf
        	end if

            If pblnUseSubCategory Then
        	If prssubCategory.RecordCount > 0 Then
        		.Write"<TR><TD COLSPAN=4>"
				If (Trim(prsCategory("CatID")) = plngCatID) and (len(mlngsubCatID) <> 0) Then
					.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
						 & "bgcolor='whitesmoke' style='cursor:hand;' id='tbl" & prsCategory("CatID") & "'" _
						 & ">"
				Else
					.Write "<table class='tbl' width='100%' cellpadding='0' cellspacing='0' border='0' rules='none' " _
						 & "bgcolor='whitesmoke' style='display: none; cursor:hand;' id='tbl" & prsCategory("CatID") & "'" _
						 & ">"
				End If
				.Write "<colgroup align='left' width='5%'>"
				.Write "<colgroup align='left' width='25%'>"
				.Write "<colgroup align='left' width='52%'>"
				.Write "<colgroup align='left' width='13%'>"
	        	do while not prssubCategory.EOF
					pstrURL = "sfCategoryAdmin.asp?Action=View&subCatID=" & prssubCategory("subCatID")
					pstrTitle = "Click to view " & prssubCategory("subcatName")
					
					If (isNull(prssubCategory("subCatIsActive")) or len(prssubCategory("subCatIsActive"))=0) Then
						pblnActive = false
					Else
						pblnActive = cBool(prssubCategory("subCatIsActive"))
					End If
					
					if pblnActive then
						If Trim(prsCategory("CatID")) = plngCatID Then
							.Write " <TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "SelectCategory('" & prssubCategory("subcatID") & "')" & chr(34) & ">"
							.Write " <TD>&nbsp;</TD>"
							.Write " <TD>&nbsp;&nbsp;&nbsp;" & prssubCategory("subcatName") & "</TD>"
						Else
							.Write " <TR class='Active' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
							.Write " <TD>&nbsp;</TD>"
       						.Write "<TD>&nbsp;&nbsp;&nbsp;<a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & prssubCategory("subcatName") & "&nbsp;</a></TD>"
						End If
					else
						If Trim(prsCategory("CatID")) = plngCatID Then
							.Write " <TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "SelectCategory('" & prssubCategory("subcatID") & "')" & chr(34) & ">"
							.Write " <TD>&nbsp;</TD>"
							.Write " <TD>&nbsp;&nbsp;&nbsp;" & prssubCategory("subcatName") & "</TD>"
						Else
							.Write " <TR class='Inactive' title='" & pstrTitle & "' onmouseover='doMouseOverRow(this); DisplayTitle(this);' onmouseout='doMouseOutRow(this); ClearTitle();' onmousedown=" & chr(34) & "doMouseClickRow('" & pstrURL & "')" & chr(34) & ">"
							.Write " <TD>&nbsp;</TD>"
       						.Write "<TD>&nbsp;&nbsp;&nbsp;<a href='" & pstrURL & "' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='" & pstrTitle & "'>" & prssubCategory("subcatName") & "&nbsp;</a></TD>"
						End If
					End If
					
       				.Write "<TD>&nbsp;&nbsp;&nbsp;" & prssubCategory("subcatDescription") & "&nbsp;</TD>"
					if cBool(prssubCategory("subCatIsActive")) then
        				.Write "<TD><a href='sfCategoryAdmin.asp?Action=Deactivate&subCatID=" & prssubCategory("subCatID") & _
												"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to deactivate " & prssubCategory("subcatName") & ".'>Active</a></TD></TR>" & vbCrLf
					else
        				.Write "<TD><a href='sfCategoryAdmin.asp?Action=Activate&subCatID=" & prssubCategory("subCatID") & _
												"' onMouseOver='return DisplayTitle(this);' onMouseOut='ClearTitle();' title='Click to activate " & prssubCategory("subcatName") & ".'>Inactive</a></TD></TR>" & vbCrLf
        			end if
   					prssubCategory.MoveNext
        		loop
        		.Write "</TABLE>"
        		.Write "</TD></TR>"
        	End If
        	End If
            prsCategory.MoveNext
        Next
    Else
			.Write "<TR><TD align=center><h3>There are no Categories</h3></TD></TR>"
    End If
		.Write "</td></tr></TABLE></div>"
		.Write "</TABLE>"
	End With
	
End Sub      'OutputSummary

'***********************************************************************************************

Function ValidateValues()

Dim strError

    strError = ""

    If Len(pstrCatName) = 0 Then
        strError = strError & "Please enter a category name." & cstrDelimeter
    End If

    pstrMessage = strError
    ValidateValues = (Len(strError) = 0)


End Function 'ValidateValues

End Class   'clsCategory

Function SetImagePath(strImage)

	If len(trim(strImage)) > 0 Then
		SetImagePath = mstrBaseHRef & strImage
	Else
		SetImagePath = "images/NoImage.gif"
	End If

End Function

Sub LoadFilter

dim pstrOrderBy

	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")
	mblnShowSummary = (lCase(trim(Request.Form("blnShowSummary"))) = "false")

'Build Filter

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'Category
			pstrOrderBy = "catName"
		Case "2"	'Description
			pstrOrderBy = "catDescription"
		Case "3"	'Active
			pstrOrderBy = "catIsActive"
	End Select	

	mstrsqlWhere = " Order By " & pstrOrderBy & " " & mstrSortOrder
	
End Sub    'LoadFilter

mstrPageTitle = "Category Administration"

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
'Assumptions:
'   cnn: defines a previously opened connection to the database

'page variables
Dim mAction
Dim mclsCategory
Dim mlngCatID, mlngsubCatID
Dim mblnShowSummary
Dim mstrsqlWhere, mstrSortOrder,mstrOrderBy

    mlngCatID = Request.QueryString("CatID")
    If len(mlngCatID) = 0 Then mlngCatID = Request.Form("CatID")

    mlngsubCatID = Request.QueryString("subCatID")
    If len(mlngsubCatID) = 0 Then mlngsubCatID = Request.Form("subCatID")

    mAction = Request.QueryString("Action")
    If Len(mAction) = 0 Then mAction = Request.Form("Action")
    
    Call LoadFilter
    
    Set mclsCategory = New clsCategory
    With mclsCategory
    
    Select Case mAction
        Case "New", "Update"
            .Update
            If .Load Then .FindCategory mlngCatID
        Case "DeleteSub"
			.DeletesubCategory mlngsubCatID
			If .Load Then .FindCategory mlngCatID
        Case "DeleteCat"
			.DeleteCategory mlngCatID
            .Load
        Case "View"
			If len(mlngCatID) > 0 Then
				If .Load Then .FindCategory mlngCatID
			Else
				If .Load Then .FindsubCategory mlngsubCatID
			End If
        Case "Activate", "Deactivate"
			If len(mlngCatID) > 0 Then
				.ActivateCategory mlngCatID, mAction = "Activate"
				If .Load Then .FindCategory mlngCatID
			Else
				.ActivatesubCategory mlngsubCatID, mAction = "Activate"
				If .Load Then .FindsubCategory mlngsubCatID
			End If
		Case "CopyCat"
			.CopyCategory mlngCatID
			If .Load Then .FindCategory mlngCatID
        Case Else
            .Load
    End Select
    
Call WriteHeader("body_onload();",True)
%>

<SCRIPT LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
var theDataForm;
var theKeyField;
var strDetailTitle = "<%= .catName %> Details";
var strsubDetailTitle = "<%= .subcatName %> Details";
var pblnAddSub;

function MakeSubDirty(theItem)
{
var theForm = theItem.form;

	pblnAddSub = true;
	theForm.btnUpdate.disabled = false;
	theForm.btnReset.disabled = false;
}

function MakeCatDirty(theItem)
{
var theForm = theItem.form;

	theForm.btnUpdate.disabled = false;
	theForm.btnReset.disabled = false;
}

function body_onload()
{
	theDataForm = document.frmData;
	theKeyField = theDataForm.affID;
	pblnAddSub = false;
<% If mblnShowSummary then Response.Write "DisplaySummary();" & vbcrlf %>
}

var gobjImage;
var gblnSwitch;

function SelectImage(theImage)
{
	gblnSwitch = true;
	gobjImage = theImage;
	document.frmData.tempFile.click();
	return false;
}

function ProcessPath(theFile)
{
var pstrFilePath = theFile.value;
var pstrBaseHRef = document.frmData.strBaseHRef.value;
var pstrBasePath = document.frmData.strBasePath.value;
var pstrHREF;
var pstrItem;

	if (gblnSwitch)
	{
	gobjImage.src = pstrFilePath;
	pstrItem = gobjImage.name.replace("img","");;
	pstrHREF = pstrFilePath.replace(pstrBasePath,"");
	eval("document.frmData." + pstrItem).value = pstrHREF;
	document.frmData.btnReset.disabled = false;
	document.frmData.btnUpdate.disabled = false;
	gblnSwitch = false;
	theFile.value = "";
	}
}

function DuplicateCategory(theButton)
{
var theForm = theButton.form;
var pstrNewCatName = prompt("Enter New Category Name","New Category Name");

if (pstrNewCatName != null)
{
    theForm.CopyCat.value = pstrNewCatName;
    theForm.Action.value = "CopyCat";
    theForm.submit();
}
}

function btnNewCategory_onclick(theButton)
{
var theForm = theButton.form;

	pblnAddSub = false;

    theForm.catID.value = "";
    theForm.catName.value = "";
    theForm.catDescription.value = "";
    theForm.catImage.value = "";
    theForm.catIsActive.checked = false;
    theForm.catHttp.value = "";

    theForm.btnUpdate.value = "Add Category";
    theForm.btnDeleteCat.disabled = true;
    theForm.btnDeleteSub.disabled = true;
//	theForm.btnUpdate.disabled = true;
	theForm.btnReset.disabled = true;
    theForm.catName.focus();
    document.all("spancatName").innerHTML = theDataForm.btnUpdate.value;

	theForm.subcatID.value = "";
	theForm.subcatName.value = "";
	theForm.subcatDescription.value = "";
	theForm.subcatImage.value = "";
	theForm.subcatIsActive.checked = false;
	theForm.btnDeleteSub.disabled = true;
	theForm.btnReset.disabled = false;
    document.all("spansubcatName").innerHTML = "&nbsp;";
}

function btnNewSubCategory_onclick(theButton)
{
var theForm = theButton.form;

	pblnAddSub = true;

	theForm.subcatName.value = "";
	theForm.subcatDescription.value = "";
	theForm.subcatImage.value = "";
	theForm.subcatIsActive.checked = false;
    theForm.btnDeleteSub.disabled = true;
    theForm.btnUpdate.value = "Add sub-Category";
    document.all("spansubcatName").innerHTML = theDataForm.btnUpdate.value;
    theForm.subcatName.focus();

}

function btnDeleteCat_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.catName.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "DeleteCat";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function btnDeleteSub_onclick(theButton)
{
var theForm = theButton.form;
var blnConfirm;

    blnConfirm = confirm("Are you sure you wish to delete " + document.frmData.subcatName.value + "?");
    if (blnConfirm)
    {
    theForm.Action.value = "DeleteSub";
    theForm.submit();
    return(true);
    }
    else
    {
    return(false);
    }
}

function btnReset_onclick(theButton)
{
var theForm = theButton.form;

	pblnAddSub = false;
    theForm.Action.value = "Update";
    theForm.btnUpdate.value = "Save Changes";
    document.all("spancatName").innerHTML = strDetailTitle;
    document.all("spansubcatName").innerHTML = strsubDetailTitle;
    document.imgcatImage.src = "<%= SetImagePath(.catImage) %>";
    document.imgsubcatImage.src = "<%= SetImagePath(.subcatImage) %>";
//    theForm.btnUpdate.disabled = true;
    theForm.btnDeleteSub.disabled = false;
    theForm.btnDeleteCat.disabled = false;
}

function DisplaySection(strSection)
{
var pstrSection = "tbl" + strSection;

  frmData.Show.value = strSection;
  HideSections();
  document.all(pstrSection).style.display = "";

return(false);
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

function SelectCategory(lngsubCatID)
{
	theDataForm.subcatID.value = lngsubCatID;
	ChangeSubCategory(theDataForm.subcatID);
}

function ChangeSubCategory(theSelect)
{
var theForm = theSelect.form;
var intIndex = theSelect.selectedIndex;

<% .OutputSubCatValues %>

	if (intIndex == 0)
	{
	btnNewSubCategory_onclick(theForm.subcatID);
	}
	else
	{
	theForm.subcatName.value = arysubcatName[intIndex];
	theForm.subcatDescription.value = arysubcatDescription[intIndex];
	theForm.subcatImage.value = arysubcatImage[intIndex];
	theForm.subcatIsActive.checked = arysubcatIsActive[intIndex];
	theForm.btnDeleteSub.disabled = false;
    document.all("spansubcatName").innerHTML = arysubcatName[intIndex] + " Details";
	}
}

function ValidInput(theForm)
{
  if (theDataForm.catName.value == "")
  {
    alert("Please enter a category name.")
    theDataForm.catName.focus();
    return(false);
  }
	
  if (pblnAddSub)
  {
	if (theDataForm.subcatName.value == "")
	{
	  alert("Please enter a sub-category name.")
	  theDataForm.subcatName.focus();
	  return(false);
	}
  }
	
    return(true);
}

function ExpandCategory(theLink,lngID)
{

	if (theLink.innerHTML == "+")
	{
		theLink.innerHTML = "-"
		theLink.title = "Hide sub-categories"
		eval("tbl" + lngID).style.display = "";
	}
	else
	{
		theLink.innerHTML = "+"
		theLink.title = "Show sub-categories"
		eval("tbl" + lngID).style.display = "none";
	}
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

<FORM action='sfCategoryAdmin.asp' id=frmData name=frmData onsubmit='return ValidInput();' method=post>
<input type=hidden id=catID name=catID value=<%= .catID %>>
<input type=hidden id=CopyCat name=CopyCat>
<input type=hidden id=Action name=Action value='Update'>
<input type=hidden id=blnShowSummary name=blnShowSummary value=''>
<input type=hidden id=OrderBy name=OrderBy value='<%= mstrOrderBy %>'>
<input type=hidden id=SortOrder name=SortOrder value='<%= mstrSortOrder %>'>

<INPUT type=hidden id=strBaseHRef name=strBaseHRef Value='<%= mstrBaseHRef %>'>
<INPUT type=hidden id=strBasePath name=strBasePath Value='<%= mstrBasePath %>'>

<%= .OutputSummary %>

<span id=spantempFile style="display:none">
<input type=file id=tempFile name=tempFile onchange="ProcessPath(this);">
</span>
<table class="tbl" width="95%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <colgroup align=right>
  <colgroup align=left>
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spancatName"><%= .catName %> Details</span></th>
  </tr>
      <TR>
        <TD>Name:</TD>
        <TD><INPUT id=catName onchange='MakeCatDirty(this);' name=catName Value='<%= .catName %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>Description:</TD>
        <TD><INPUT id=catDescription onchange='MakeCatDirty(this);' name=catDescription Value='<%= .catDescription %>' maxlength=255 size=30></TD>
      </TR>
       <TR>
        <TD class="Label">Image:</TD>
        <TD><INPUT id=catImage onchange='MakeCatDirty(this);' name=catImage Value='<%= .catImage %>' maxlength=200 size=60>
			<img style="cursor:hand" name=imgcatImage id=imgcatImage border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"return ClearTitle();" src="<%= SetImagePath(.catImage) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit the category image">
		</TD>
      </TR>
      <TR>
        <TD>URL:</TD>
        <TD><INPUT id=catHttp onchange='MakeCatDirty(this);' name=catHttp Value='<%= .catHttp %>' maxlength=255 size=60></TD>
      </TR>
      <TR>
        <TD>&nbsp;</TD>
        <TD><INPUT type=checkbox id=catIsActive onchange='MakeCatDirty(this);' name=catIsActive <% If .CatIsActive Then Response.Write "Checked" %>>&nbsp;Is Active</TD>
      </TR>
  <% If .UseSubCategory Then %>
  <tr class='tblhdr'>
	<th colspan="2" align=center><span id="spansubcatName"><%= .subcatName %> Details</span></th>
  </tr>
      <TR>
        <TD>&nbsp;</TD>
        <TD>
			<select size="1"  id=subcatID name=subcatID onchange="ChangeSubCategory(this);">
			<option value="">Create New sub-Category</option>
			<% 
				If isObject(.rssubCategory) Then 
					.rssubCategory.Filter = "subcatCategoryID=" & .CatID
					if .rssubCategory.recordcount > 0 Then .rssubCategory.movefirst
					Call MakeCombo(.rssubCategory,"subCatName","subCatID",.subcatID)
				End If
			%>
			</select>
		</TD>        
      </TR>
      <TR>
        <TD>Name:</TD>
        <TD><INPUT id=subcatName onchange='MakeSubDirty(this)' name=subcatName Value='<%= .subcatName %>' maxlength=30 size=30></TD>
      </TR>
      <TR>
        <TD>Description:</TD>
        <TD><INPUT id=subcatDescription onchange='MakeSubDirty(this)' name=subcatDescription Value='<%= .subcatDescription %>' maxlength=255 size=30></TD>
      </TR>
       <TR>
        <TD class="Label">Image:</TD>
        <TD><INPUT id=subcatImage onchange='MakeSubDirty(this)' name=subcatImage Value='<%= .subcatImage %>' maxlength=200 size=60>
			<img style="cursor:hand" name=imgsubcatImage id=imgsubcatImage border="0" 
				 onmouseover="DisplayTitle(this);" onmouseout"ClearTitle();" src="<%= SetImagePath(.subcatImage) %>" 
				 onclick="return SelectImage(this);" 
				 title="Click to edit the sub-category image">
		</TD>
      </TR>
      <TR>
        <TD>&nbsp;</TD>
        <TD><INPUT type=checkbox id=subcatIsActive onchange='MakeSubDirty(this)' name=subcatIsActive <% If (.subcatIsActive) Then Response.Write "Checked" %>>&nbsp;Is Active</TD>
      </TR>
  <% End If %>

  <TR>
    <TD>&nbsp;</TD>
    <TD>
        <INPUT class='butn' id=btnNewCat name=btnNewCat type=button value='New Category' onclick='return btnNewCategory_onclick(this)'>&nbsp;
        <INPUT class='butn' id=btnCopyCat name=btnCopyCat type=button value='Copy Category' onclick='return DuplicateCategory(this)'>&nbsp;
        <INPUT class='butn' id=btnReset name=btnReset type=reset value=Reset onclick='return btnReset_onclick(this)' disabled>&nbsp;&nbsp;
        <INPUT class='butn' id=btnDeleteCat name=btnDeleteCat type=button value='Delete Category' onclick='return btnDeleteCat_onclick(this)'>
        <% If .UseSubCategory Then %><INPUT class='butn' id=btnDeleteSub name=btnDeleteSub type=button value='Delete sub-Category' onclick='return btnDeleteSub_onclick(this)' <% If Not(.catHasSubCategory And len(.subCatID)<>0) Then Response.Write "disabled"  %>><% End If %>
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
    Set mclsCategory = Nothing
    Set cnn = Nothing
    Response.Flush
%>
