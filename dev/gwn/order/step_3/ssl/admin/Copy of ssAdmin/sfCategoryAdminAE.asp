<%
'********************************************************************************
'*   Category Manager Version SF 5.0 AE                                         *
'*   Release Version:	2.00.003		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		April 9, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   2.00.004B (May 29, 2005)							                        *
'*   - Modified core to use XML document to build category structure		    *
'*                                                                              *
'*   2.00.003 (April 9, 2004)							                        *
'*   - Bug Fix: Resolved unable to delete top level category when 1st entry     *
'*                                                                              *
'*   2.00.002 (March 11, 2004)							                        *
'*   - Added support for HTML tags in description						        *
'*   - Added support for image selector									        *
'*   - Added clean-up code to correct for common category db errors		        *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************
Option Explicit
Response.Buffer = True
%>
<!--#include file="SSLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'********************************************************************************************
' Main Execution
'********************************************************************************************

Const cbytMode = 1	'for internal use: 2-full view, 1-partial view, 0-name only (SF default)
Const cblnDebug = False	'True	'False

Dim mstrAction
Dim mstrcatID 
Dim mstrcatParent
Dim mbytDepth
Dim mstrcatName
Dim mstrcatDescription
Dim mstrcatURL
Dim mstrcatImage
Dim mblncatIsActive
Dim mstrStatus
Dim mstrCatHeir
Dim mstrcatBottom
Dim mstrResult
    
mstrAction = Request.QueryString("Action")
If Len(mstrAction) = 0 Then mstrAction = Request.Form("Action")

If cblnDebug Then Call InitializeDebugging("<path>\CatMgrDebug.txt", True)
DebugPrintLater "mstrAction", mstrAction
Select Case mstrAction
	Case "GetData"
		If True Then	'True	False
			Response.Write AECategories_XML
		Else
			Response.Write "<?xml version=""1.0"" encoding=""utf-16""?>" & vbcrlf
			Response.Write "<root>" & vbcrlf
			Response.Write "<categories>" & vbcrlf
			Response.Write AECategories
			Response.Write "</categories>" & vbcrlf
			Response.Write "</root>" & vbcrlf
		End If
	Case "UpdateData"
		mstrResult = ProcessNodes
		Call OutputHTMLPage
	Case "SaveChanges"
		mstrResult = ProcessNodes
	Case "CleanCategories"
		mstrResult = ProcessNodes
		Response.Write mstrResult
	Case Else
		Call CheckCategoryCleanliness
		Call OutputHTMLPage
End Select

'added for debugging
If cblnDebug Then Call CleanUpDebugging("")

cnn.Close
Set cnn = Nothing
Set debug = Nothing

'********************************************************************************************

Function CheckCategoryCleanliness

Dim pstrSQL
Dim pobjRS
Dim pblnNeedsCleaning: pblnNeedsCleaning = False

	pstrSQL = "SELECT sfSub_Categories.subcatID" _
			& " FROM sfCategories RIGHT JOIN sfSub_Categories ON sfCategories.catID = sfSub_Categories.subcatCategoryId" _
			& " WHERE sfCategories.catID Is Null"
	Set pobjRS = GetRS(pstrSQL)
	If Not pobjRS.EOF Then 
		Response.Write "<h2>There are unmatched subcategories in the database.</h2>"
		pblnNeedsCleaning = True
	End If
	Set pobjRS = Nothing

	pstrSQL = "SELECT catID from sfCategories Where catID Not In (SELECT sfSub_Categories.subcatCategoryId" _
			& " FROM sfSub_Categories" _
			& " WHERE sfSub_Categories.Depth=0)"
	Set pobjRS = GetRS(pstrSQL)
	If Not pobjRS.EOF Then 
		Response.Write "<h2>There are categories without required subcategories in the database.</h2>"
		pblnNeedsCleaning = True
	End If
	Set pobjRS = Nothing
	
	If pblnNeedsCleaning Then
		Call CleanCategories
	End If

End Function	'CheckCategoryCleanliness

'********************************************************************************************

Function CleanCategories

Dim pstrSQL
Dim pobjRS
Dim pobjRSSubCategory
Dim pblnNeedsCleaning: pblnNeedsCleaning = False

	pstrSQL = "SELECT subcatID" _
			& " FROM sfCategories RIGHT JOIN sfSub_Categories ON sfCategories.catID = sfSub_Categories.subcatCategoryId" _
			& " WHERE sfCategories.catID Is Null"
			
	pstrSQL = "Delete From sfSub_Categories Where subcatID in (" & pstrSQL & ")"
	cnn.Execute pstrSQL,,128
	
	Response.Write "<h2>Unmatched subcategories deleted.</h2>"

	pstrSQL = "SELECT * from sfCategories Where catID Not In (SELECT sfSub_Categories.subcatCategoryId" _
			& " FROM sfSub_Categories" _
			& " WHERE sfSub_Categories.Depth=0)"
	Set pobjRS = GetRS(pstrSQL)
	
	If Not pobjRS.EOF Then
		Do While Not pobjRS.EOF 
			Call CreatePrimarySubCategory(pobjRS.Fields("catID").Value, pobjRS.Fields("catName").Value)
			Response.Write "<h2>Category: " & pobjRS.Fields("catID").Value & " - " & pobjRS.Fields("catName").Value & " updated.</h2>"
			pobjRS.Movenext
		Loop
	End If	'Not pobjRS.EOF
	Set pobjRS = Nothing

End Function	'CleanCategories

'********************************************************************************************

Function CreatePrimarySubCategory(byVal lngCatID, byVal strCatName)

Dim pstrSQL
Dim pobjRS

	pstrSQL = "INSERT INTO sfSub_Categories (subcatCategoryId, subcatName, Depth, HasProds, bottom, CatHierarchy)" _
			& " VALUES (" _
			& wrapSQLValue(lngCatID, True, enDatatype_string) & "," _
			& wrapSQLValue(strCatName, True, enDatatype_string) & "," _
			& "0," _
			& "0," _
			& "0," _
			& wrapSQLValue(Session.SessionID, False, enDatatype_string) _
			& ")"
	'debugprint "pstrSQL", pstrSQL
	cnn.Execute pstrSQL,,128
	
	pstrSQL = "Select subcatID From sfSub_Categories Where CatHierarchy=" & wrapSQLValue(Session.SessionID, False, enDatatype_string)
	'debugprint "pstrSQL", pstrSQL
	Set pobjRS = GetRS(pstrSQL)
	
	If Not pobjRS.EOF Then
		pstrSQL = "Update sfSub_Categories Set CatHierarchy=" & wrapSQLValue("none-" & pobjRS.Fields("subcatID").Value, False, enDatatype_string) & " Where subCatID=" & wrapSQLValue(pobjRS.Fields("subcatID").Value, False, enDatatype_number)
		'debugprint "pstrSQL", pstrSQL
		cnn.Execute pstrSQL,,128
	End If	'Not pobjRS.EOF
	Set pobjRS = Nothing
			
End Function	'CreatePrimarySubCategory

'********************************************************************************************
' Update Data Section
'********************************************************************************************


	Function ProcessNodes

	Dim nodeList
	Dim pobjNode
	Dim pstrResult
	Dim i,j,k
	Dim mobjXMLDoc
	Dim plngNewCatID
	Dim plngNewSubCatID
	Dim pstrTempXML
	Dim pobjRSCategory
	Dim pobjRSSubCategory
	Dim plngID
	Dim plngTempID
	Dim plngParentID
	Dim pstrXMLData
	
	On Error Resume Next
	
		set mobjXMLDoc = CreateObject("MSXML2.DOMDocument")
		mobjXMLDoc.async = false

		pstrXMLData = Request.Form("xmlData")
		If Len(pstrXMLData) = 0 Then pstrXMLData = Request.Form
		DebugPrintLater "pstrXMLData", pstrXMLData
		'debug.printlater "pstrXMLData", pstrXMLData
		'debug.WriteToFile False, "D:\Sandshot Software\Website\Add-Ons 4 StoreFront\StoreFront-5AE-Demo\ssl\admin\SSAdmin\SSLibrary\categoryDebug.txt", False
		
		If mobjXMLDoc.loadXML(pstrXMLData) Then
		
			Set nodeList = mobjXMLDoc.getElementsByTagName("category")

			Set pobjRSCategory = server.CreateObject("adodb.Recordset")
			pobjRSCategory.CursorLocation = 3
			pobjRSCategory.Open "sfCategories", cnn, 1, 3, 2

			Set pobjRSSubCategory = server.CreateObject("adodb.Recordset")
			pobjRSSubCategory.CursorLocation = 3
			pobjRSSubCategory.open "sfSub_Categories", cnn, 1, 3, 2
			
			'First process the Inserts to categories to get IDs
			For i =0 To nodeList.length - 1
				Call GetNodeValues(nodeList.Item(i))

				If mstrStatus = "Insert" Then
				
					If Len(mstrCatHeir) = 0 Then	'this is a category
					
						With pobjRSCategory
							.AddNew
							.Fields("catName").Value = mstrcatName
							If cbytMode > 0 Then .Fields("catDescription").Value = mstrcatDescription
							If cbytMode > 0 Then .Fields("catHttpAdd").Value = mstrcatURL
							If cbytMode > 0 Then .Fields("catImage").Value = mstrcatImage
							.Fields("catIsActive").Value = Abs(mblncatIsActive)
							.Fields("catHasSubCategory").Value = Abs(mstrcatBottom)
							.Update
							plngNewCatID = .Fields("catID").Value
							
						End With
						
						'now insert the duplicate
						With pobjRSSubCategory
							.AddNew
							.Fields("subcatCategoryId").Value = plngNewCatID
							.Fields("subcatName").Value = mstrcatName
							If cbytMode > 0 Then .Fields("subcatDescription").Value = mstrcatDescription
							If cbytMode > 0 Then .Fields("subcatHttpAdd").Value = mstrcatURL
							If Len(mstrcatImage) > 0 And cbytMode > 0 Then .Fields("subcatImage").Value = mstrcatImage
							.Fields("subcatIsActive").Value = Abs(mblncatIsActive)
							.Fields("Depth").Value = 0
							.Fields("HasProds").Value = 0
							.Fields("bottom").Value = mstrcatBottom
							.Update
							plngNewSubCatID = .Fields("subcatID").Value

							.Fields("CatHierarchy").Value = "none-" & plngNewSubCatID
							.Update
							
						End With
						
						pstrTempXML = mobjXMLDoc.xml
						set mobjXMLDoc = CreateObject("MSXML2.DOMDocument")
						mobjXMLDoc.async = false
							
						'need to replace parent with plngNewCatID
						pstrTempXML = Replace(pstrTempXML,"<catParent>" & mstrcatID & "</catParent>","<catParent>" & plngNewCatID & "</catParent>")

						'need to replace hierarchy with plngNewSubCatID
						pstrTempXML = Replace(pstrTempXML,mstrcatID,plngNewCatID)

						mobjXMLDoc.loadXML pstrTempXML
					
						Set nodeList = mobjXMLDoc.getElementsByTagName("category")
							
					End If	'this is a category

				End If	'this is an insert
			Next 'i

			'Now process the Inserts for the sub-categories to get IDs
			For i =0 To nodeList.length - 1
				Call GetNodeValues(nodeList.Item(i))
				If mstrStatus = "Insert" Then
				
					If Len(mstrCatHeir) > 0 Then	'this is a sub-category
					
						plngParentID = Replace(mstrcatParent,"root:","")
						plngParentID = Replace(plngParentID,"sub:","")
						With pobjRSSubCategory
							.AddNew
							.Fields("subcatCategoryId").Value = CLng(plngParentID)
							.Fields("subcatName").Value = mstrcatName
							If cbytMode > 0 Then .Fields("subcatDescription").Value = mstrcatDescription
							If cbytMode > 0 Then .Fields("subcatHttpAdd").Value = mstrcatURL
							If Len(mstrcatImage) > 0 And cbytMode > 0 Then .Fields("subcatImage").Value = mstrcatImage
							.Fields("subcatIsActive").Value = Abs(mblncatIsActive)
							.Fields("Depth").Value = mbytDepth
							.Fields("HasProds").Value = 0
							.Fields("bottom").Value = mstrcatBottom
							.Update
							plngNewSubCatID = .Fields("subcatID").Value
							
							.Fields("CatHierarchy").Value = Replace(mstrCatHeir,mstrcatID,plngNewSubCatID)
							.Update
							
						End With
						
						pstrTempXML = mobjXMLDoc.xml
						set mobjXMLDoc = CreateObject("MSXML2.DOMDocument")
						mobjXMLDoc.async = false
							
						'need to replace hierarchy with plngNewSubCatID
						pstrTempXML = Replace(pstrTempXML,mstrcatID,plngNewSubCatID)

						mobjXMLDoc.loadXML pstrTempXML
						
						Set nodeList = mobjXMLDoc.getElementsByTagName("category")
							
					End If	'this is a sub-category

				End If
			Next 'i
			
			pstrTempXML = mobjXMLDoc.xml
			set mobjXMLDoc = CreateObject("MSXML2.DOMDocument")
			mobjXMLDoc.async = false
							
			'need to change inserts to updates
			pstrTempXML = Replace(pstrTempXML,"<catStatus>Insert</catStatus>","<catStatus>Update</catStatus>")

			mobjXMLDoc.loadXML pstrTempXML
			Set nodeList = mobjXMLDoc.getElementsByTagName("category")

			'now do the rest
			For i =0 To nodeList.length - 1

				Call GetNodeValues(nodeList.Item(i))

				plngTempID = Replace(mstrcatID,"root:","")
				plngTempID = Replace(plngTempID,"sub:","")

				Select Case mstrStatus
					Case "Delete"
						If Instr(1,mstrcatID,"root:") <> 0 Then
							cnn.Execute "Delete from sfCategories where catID = " & plngTempID,,128
						ElseIf Instr(1,mstrcatID,"sub:") <> 0 Then
							cnn.Execute "Delete from sfSub_Categories where subcatID = " & plngTempID,,128
							cnn.Execute "Delete from sfSubCatDetail where subcatCategoryId = " & plngTempID,,128
						End If
					Case "Update"
						If Len(mstrCatHeir) = 0 Then	'this is a category
					
							If Instr(1,plngTempID,"sub:") <> 0 Then
							'this was a sub-category moved to the root
							'so add a category, get the id, then replace all sub-categories (including this one) with the new ID

								With pobjRSCategory
									.AddNew
									.Fields("catName").Value = mstrcatName
									If cbytMode > 0 Then .Fields("catDescription").Value = mstrcatDescription
									If cbytMode > 0 Then .Fields("catHttpAdd").Value = mstrcatURL
									If cbytMode > 0 Then .Fields("catImage").Value = mstrcatImage
									.Fields("catIsActive").Value = Abs(mblncatIsActive)
									.Fields("catHasSubCategory").Value = Abs(mstrcatBottom)
									.Update
									plngNewCatID = .Fields("catID").Value
									
								End With
								
								With pobjRSSubCategory
									.Filter = "subcatID = " & plngTempID

									.Fields("subcatCategoryId").Value = CLng(mstrcatParent)
									.Fields("subcatName").Value = mstrcatName
									If cbytMode > 0 Then .Fields("subcatDescription").Value = mstrcatDescription
									If cbytMode > 0 Then .Fields("subcatHttpAdd").Value = mstrcatURL
									If Len(mstrcatImage) > 0 And cbytMode > 0 Then .Fields("subcatImage").Value = mstrcatImage
									.Fields("subcatIsActive").Value = Abs(mblncatIsActive)
									.Fields("Depth").Value = mbytDepth
									'.Fields("HasProds").Value = 0
									.Fields("bottom").Value = mstrcatBottom
									.Fields("CatHierarchy").Value = mstrCatHeir
									.Update
								End With

								pstrTempXML = mobjXMLDoc.xml
								set mobjXMLDoc = CreateObject("MSXML2.DOMDocument")
								mobjXMLDoc.async = false
									
								'need to replace parent with plngNewCatID
								pstrTempXML = Replace(pstrTempXML,"<catParent>" & plngNewCatID & "</catParent>","<catParent>" & plngNewCatID & "</catParent>")

								mobjXMLDoc.loadXML pstrTempXML
					
								Set nodeList = mobjXMLDoc.getElementsByTagName("category")

							Else
							'this is a category to be updated
							
								With pobjRSCategory
									.Filter = "catID = " & plngTempID
									.Fields("catName").Value = mstrcatName
									If cbytMode > 0 Then .Fields("catDescription").Value = mstrcatDescription
									If cbytMode > 0 Then .Fields("catHttpAdd").Value = mstrcatURL
									If cbytMode > 0 Then .Fields("catImage").Value = mstrcatImage
									.Fields("catIsActive").Value = Abs(mblncatIsActive)
									.Fields("catHasSubCategory").Value = Abs(mstrcatBottom)
									.Update
								End With

								
								'Now update the corresponding subCategory
								With pobjRSSubCategory
									.Filter = "CatHierarchy like 'none-%' AND subcatCategoryId = " & plngTempID

									.Fields("subcatName").Value = mstrcatName
									If cbytMode > 0 Then .Fields("subcatDescription").Value = mstrcatDescription
									If cbytMode > 0 Then .Fields("subcatHttpAdd").Value = mstrcatURL
									If Len(mstrcatImage) > 0 And cbytMode > 0 Then .Fields("subcatImage").Value = mstrcatImage
									.Fields("subcatIsActive").Value = Abs(mblncatIsActive)
									.Fields("Depth").Value = mbytDepth
									.Fields("bottom").Value = mstrcatBottom
									.Update
								End With

							End If
						
						Else	'this is a sub-category
					
							If Instr(1,plngTempID,"root:") = 0 Then
							'this is a sub-category to be updated
							
								With pobjRSSubCategory
									.Filter = "subcatID = " & plngTempID

									.Fields("subcatCategoryId").Value = CLng(mstrcatParent)
									.Fields("subcatName").Value = mstrcatName
									If cbytMode > 0 Then .Fields("subcatDescription").Value = mstrcatDescription
									If cbytMode > 0 Then .Fields("subcatHttpAdd").Value = mstrcatURL
									If Len(mstrcatImage) > 0 And cbytMode > 0 Then .Fields("subcatImage").Value = mstrcatImage
									.Fields("subcatIsActive").Value = Abs(mblncatIsActive)
									.Fields("Depth").Value = mbytDepth
									'.Fields("HasProds").Value = 0
									.Fields("bottom").Value = mstrcatBottom
									.Fields("CatHierarchy").Value = mstrCatHeir
									.Update
								End With

							Else
							'this was a category moved down so delete

								pstrSQL = "Delete from sfCategories where catID=" & plngTempID
								cnn.Execute pstrSQL,,128

							End If
						End If
				End Select
			Next 'i
			
			'Now need to update the sfSubCatDetail table to put all unassigned items into "No Category"
			Dim pstrSQL
			Dim plngNoCategoryID
			Dim pobjNoCategories
			
			With pobjRSSubCategory
				.Filter = "subcatName = 'No Category'"
				
				If .EOF Then
					.AddNew
					.Fields("subcatCategoryId").Value = 1
					.Fields("subcatName").Value = "No Category"
					.Fields("subcatIsActive").Value = Abs(mblncatIsActive)
					.Fields("Depth").Value = 0
					.Fields("HasProds").Value = 1
					.Fields("bottom").Value = 1
					.Update
					plngNoCategoryID = .Fields("subcatID").Value
					.Fields("CatHierarchy").Value = "none-" & plngNoCategoryID
					.Update
				Else
					plngNoCategoryID = .Fields("subcatID").Value
				End If
			End With
			
			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName" _
					& " FROM sfProducts LEFT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID" _
					& " WHERE sfSubCatDetail.ProdID Is Null"

			'Set pobjNoCategories = Server.CreateObject("ADODB.RECORDSET")
			Set pobjNoCategories = GetRS(pstrSQL)
					
			Do While Not pobjNoCategories.EOF
				pstrSQL = "Insert Into sfSubCatDetail (subcatCategoryId,ProdID,ProdName) Values (" _
						& "'" & Replace(plngNoCategoryID,"'","''") & "'," _
						& "'" & Replace(pobjNoCategories.Fields("prodID").Value,"'","''") & "'," _
						& "'" & Replace(pobjNoCategories.Fields("prodName").Value,"'","''") & "')"
				
				cnn.Execute pstrSQL,,128
				pobjNoCategories.MoveNext
			Loop
			pobjNoCategories.Close
			Set pobjNoCategories = Nothing

			Set nodeList = Nothing
			ProcessNodes = pstrResult

			pobjRSSubCategory.Close
			Set pobjRSSubCategory = Nothing
						
			pobjRSCategory.Close
			Set pobjRSCategory = Nothing
						
			Call updateCategoryHasProductsStatus
		Else
			ProcessNodes = "Error Loading XML Data"
		End If

		set mobjXMLDoc = Nothing
		Set nodeList = Nothing
		
		'Remove memory specific items
		Session.Contents.Remove("categoryCombo")	'stored for admin purposes
		Call removeFromCache("ssCategorySearch")	'site specific category menu
		
	End Function	'ProcessNodes

	'***********************************************************************************************

	Function GetXMLNodeByKey(strKey)

	Dim nodeList
	Dim i
	
		Set nodeList = mobjXMLDoc.getElementsByTagName("category")

		For i =0 To nodeList.length - 1
			If nodeList.Item(i).attributes.item(0).nodeValue = strKey Then
				Set GetXMLNodeByKey = nodeList.Item(i)
				Exit Function
			End If
		Next 'i
		
		Set nodeList = Nothing
		
	End Function	'GetXMLNodeByKey

	'********************************************************************************************

	Sub GetNodeValues(objNode)
	
	Dim i
	Dim e
	
		mstrcatID = objNode.attributes.item(0).nodeValue
		For i = 0 To objNode.childNodes.length
			Set e = objNode.childNodes.Item(i)
			Select Case e.nodeName
				Case "catParent": mstrcatParent = e.text
				Case "catDepth": mbytDepth = CInt(e.text)
				Case "catName": mstrcatName = e.text
				Case "catDescription": mstrcatDescription = e.text
				Case "catURL": mstrcatURL = e.text
				Case "catImage": mstrcatImage = e.text
				Case "catIsActive":  mblncatIsActive = e.text
				Case "catHeirarchy": mstrCatHeir = e.text
				Case "catBottom": mstrcatBottom = e.text
				Case "catStatus": mstrStatus = e.text
			End Select
		Next 'i

	End Sub	'GetNodeValues

	'********************************************************************************************

	Sub LetNodeValues(objNode)
	
	Dim i
	Dim e
	
		objNode.attributes.item(0).nodeValue = mstrcatID
		For i = 0 To objNode.childNodes.length - 1
			Set e = objNode.childNodes.Item(i)
			Select Case e.nodeName
				Case "catParent": e.text = mstrcatParent
				Case "catDepth": e.text = mbytDepth
				Case "catName": e.text = mstrcatName
				Case "catDescription": If cbytMode > 0 Then e.text = mstrcatDescription
				Case "catURL": If cbytMode > 0 Then e.text = mstrcatURL
				Case "catImage": If cbytMode > 0 Then e.text = mstrcatImage
				Case "catIsActive": e.text = mblncatIsActive
				Case "catHeirarchy": e.text = mstrCatHeir
				Case "catBottom": e.text = mstrcatBottom
				Case "catStatus": e.text = mstrStatus
			End Select
		Next 'i

	End Sub	'LetNodeValues

	'********************************************************************************************

	Sub LetElementValue(objNode, strKey, strValue)
	
	Dim i
	Dim e
	
		For i = 0 To objNode.childNodes.length - 1
			Set e = objNode.childNodes.Item(i)
			If e.nodeName = strKey Then 
				e.text = strValue
				Exit For
			End If
		Next 'i

	End Sub	'LetElementValue
	
'********************************************************************************************
' Output Category XML Section
'********************************************************************************************

'***************************************************************************************************************************************************************

Sub addNode(byRef objXMLDoc, byRef objXMLNode, byVal strNodeName, byVal strText)

Dim pobjxmlElement

	If strNodeName = "upsize_ts" Then Exit Sub

	Set pobjxmlElement = objXMLDoc.createElement(strNodeName)
	pobjxmlElement.Text = XMLEncode(strText)
	objXMLNode.appendChild pobjxmlElement

End Sub	'addNode

'***************************************************************************************************************************************************************

Sub addCDATA(byRef objXMLDoc, byRef objXMLNode, byVal strNodeName, byVal strText)

Dim pobjxmlCDATASection
Dim pobjxmlElement

	'On Error Resume Next
	
	If strNodeName = "upsize_ts" Then Exit Sub
	
	Set pobjxmlElement = objXMLDoc.createElement(strNodeName)
	Set pobjxmlCDATASection = objXMLDoc.createCDATASection(Trim(strText & ""))
	pobjxmlElement.appendChild pobjxmlCDATASection
	objXMLNode.appendChild pobjxmlElement
	
	If Err.number <> 0 Then
		Response.Write "<h2>strNodeName: " & strNodeName & "</h2>"
		Response.Write "<h2>strText: " & strText & "</h2>"
		Err.Clear
	End If

End Sub	'addCDATA

'***************************************************************************************************************************************************************

	Function AECategories
	'Purpose of this function is to return an XML document of the categories
	'
	'Note: Root Categories are set with 
	'
	'sfCategories
	'catID
	'catName
	'catDescription
	'catImage
	'catHasSubCategory
	'catIsActive
	'catHttpAdd

	'sfSub_Categories
	'subcatID
	'subcatCategoryId
	'subcatName
	'subcatDescription
	'subcatImage
	'subcatIsActive
	'CatHierarchy
	'Depth
	'bottom

	'tags
	'
	'ID				- for uniqueness root and sub are added before the id
	'Parent
	'Name
	'Description
	'Image
	'IsActive



	Dim sql
	Dim prsAEProductCategories
	Dim i
	Dim pstrOutput

	Dim pstrHeir
	Dim paHeir

		sql = "SELECT sfCategories.catID, sfCategories.catName, sfCategories.catDescription, sfCategories.catHttpAdd, sfCategories.catImage, sfCategories.catIsActive, sfCategories.catHasSubCategory " _
			& " FROM sfCategories" _
			& " ORDER BY sfCategories.catName"

		Set prsAEProductCategories = GetRS(sql)
	'	Set prsAEProductCategories = Server.CreateObject("ADODB.RECORDSET")
		With prsAEProductCategories
			For i = 1 to .RecordCount
				If (Trim(.Fields("catName").Value) <> "No Category") Then
					pstrOutput = pstrOutput & "<category ID=""" & "root:" & .Fields("catID").Value & """>" & vbcrlf
					pstrOutput = pstrOutput & MakeXMLNode("catParent","root:" & 0)
					pstrOutput = pstrOutput & MakeXMLNode("catDepth",0)
					pstrOutput = pstrOutput & MakeXMLNode("catName",.Fields("catName").Value)
					If cbytMode > 0 Then pstrOutput = pstrOutput & MakeXMLNode("catDescription",.Fields("catDescription").Value)
					If cbytMode > 0 Then pstrOutput = pstrOutput & MakeXMLNode("catURL",.Fields("catHttpAdd").Value)
					If cbytMode > 0 Then pstrOutput = pstrOutput & MakeXMLNode("catImage",.Fields("catImage").Value)
					pstrOutput = pstrOutput & MakeXMLNode("catIsActive",.Fields("catIsActive").Value)
					pstrOutput = pstrOutput & MakeXMLNode("catHeirarchy","")
					pstrOutput = pstrOutput & MakeXMLNode("catBottom",.Fields("catHasSubCategory").Value)
					pstrOutput = pstrOutput & MakeXMLNode("catStatus","")
					pstrOutput = pstrOutput & "</category>" & vbcrlf
				End If
				.MoveNext
			Next
			
			If Len(pstrOutput) = 0 Then
				pstrOutput = pstrOutput & "<category ID=""" & "new:0:" & """>" & vbcrlf
				pstrOutput = pstrOutput & MakeXMLNode("catParent","root:0")
				pstrOutput = pstrOutput & MakeXMLNode("catDepth",0)
				pstrOutput = pstrOutput & MakeXMLNode("catName","Enter you're first category")
				pstrOutput = pstrOutput & MakeXMLNode("catDescription","")
				pstrOutput = pstrOutput & MakeXMLNode("catURL","")
				pstrOutput = pstrOutput & MakeXMLNode("catImage","")
				pstrOutput = pstrOutput & MakeXMLNode("catIsActive",0)
				pstrOutput = pstrOutput & MakeXMLNode("catHeirarchy","")
				pstrOutput = pstrOutput & MakeXMLNode("catBottom",0)
				pstrOutput = pstrOutput & MakeXMLNode("catStatus","Insert")
				pstrOutput = pstrOutput & "</category>" & vbcrlf
			End If
			.Close
		End With
		Set prsAEProductCategories = Nothing
			
		'inner join used because SF doesn't always clean up after itself
		sql = "SELECT sfSub_Categories.subcatID, sfSub_Categories.subcatCategoryId, sfSub_Categories.Depth, sfSub_Categories.subcatName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatDescription, sfSub_Categories.subcatHttpAdd, sfSub_Categories.subcatImage, sfSub_Categories.bottom, sfSub_Categories.subcatIsActive" _
			& " FROM sfSub_Categories INNER JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
			& " ORDER BY sfSub_Categories.subcatCategoryId, sfSub_Categories.Depth, sfSub_Categories.subcatName;"

		Set prsAEProductCategories = GetRS(sql)
	'	Set prsAEProductCategories = Server.CreateObject("ADODB.RECORDSET")
		With prsAEProductCategories
			For i = 1 to .RecordCount
				If (Trim(.Fields("subcatName").Value) <> "No Category") Then	'Do not show No Category
	'				If Instr(1,.Fields("CatHierarchy").Value,"none-") = 0 Then	'Do not show duplicate sub-categories
	'					This test removed because information is needed downstream - it will be processed there
						pstrOutput = pstrOutput & "<category ID=""" & "sub:" & .Fields("subcatID").Value & """>" & vbcrlf
						If .Fields("Depth").Value = 0 Then
							pstrOutput = pstrOutput & MakeXMLNode("catParent","root:" & .Fields("subcatCategoryId").Value)
						ElseIf .Fields("Depth").Value = 1 Then
							pstrOutput = pstrOutput & MakeXMLNode("catParent","root:" & .Fields("subcatCategoryId").Value)
						Else
							pstrHeir = Trim(.Fields("CatHierarchy").Value)
							paHeir = Split(pstrHeir,"-")
							pstrOutput = pstrOutput & MakeXMLNode("catParent","sub:" & paHeir(.Fields("Depth").Value - 2))
						End If
						pstrOutput = pstrOutput & MakeXMLNode("catDepth",.Fields("Depth").Value)
						pstrOutput = pstrOutput & MakeXMLNode("catName",.Fields("subcatName").Value)
						If cbytMode > 0 Then pstrOutput = pstrOutput & MakeXMLNode("catDescription",.Fields("subcatDescription").Value)
						If cbytMode > 0 Then pstrOutput = pstrOutput & MakeXMLNode("catURL",.Fields("subcatHttpAdd").Value)
						If cbytMode > 0 Then pstrOutput = pstrOutput & MakeXMLNode("catImage",.Fields("subcatImage").Value)
						pstrOutput = pstrOutput & MakeXMLNode("catIsActive",.Fields("subcatIsActive").Value)
						pstrOutput = pstrOutput & MakeXMLNode("catHeirarchy",.Fields("CatHierarchy").Value)
						pstrOutput = pstrOutput & MakeXMLNode("catBottom",.Fields("bottom").Value)
						pstrOutput = pstrOutput & MakeXMLNode("catStatus","")

						pstrOutput = pstrOutput & "</category>" & vbcrlf
	'				End If
				End If
				.MoveNext
			Next
			.Close
		End With
		Set prsAEProductCategories = Nothing

		AECategories = pstrOutput
		
	End Function	'AECategories

'***************************************************************************************************************************************************************

	Function AECategories_XML
	'Purpose of this function is to return an XML document of the categories
	'
	'Note: Root Categories are set with 
	'
	'sfCategories
	'catID
	'catName
	'catDescription
	'catImage
	'catHasSubCategory
	'catIsActive
	'catHttpAdd

	'sfSub_Categories
	'subcatID
	'subcatCategoryId
	'subcatName
	'subcatDescription
	'subcatImage
	'subcatIsActive
	'CatHierarchy
	'Depth
	'bottom

	'tags
	'
	'ID				- for uniqueness root and sub are added before the id
	'Parent
	'Name
	'Description
	'Image
	'IsActive



	Dim sql
	Dim prsAEProductCategories
	Dim i
	Dim pstrOutput

	Dim pstrHeir
	Dim paHeir

		Dim xmlDoc
		Dim xmlNode
		Dim xmlRoot
		Dim xmlCategoryDetail
		
		set xmlDoc = server.CreateObject("MSXML2.DOMDocument.3.0")
		' Create processing instruction and document root
		Set xmlNode = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'")
		Set xmlNode = xmlDoc.insertBefore(xmlNode, xmlDoc.childNodes.Item(0))
	   
		' Create document root
		Set xmlRoot = xmlDoc.createElement("categories")
		Set xmlDoc.documentElement = xmlRoot
		xmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"



		sql = "SELECT sfCategories.catID, sfCategories.catName, sfCategories.catDescription, sfCategories.catHttpAdd, sfCategories.catImage, sfCategories.catIsActive, sfCategories.catHasSubCategory " _
			& " FROM sfCategories" _
			& " ORDER BY sfCategories.catName"

		Set prsAEProductCategories = GetRS(sql)
	'	Set prsAEProductCategories = Server.CreateObject("ADODB.RECORDSET")
		With prsAEProductCategories
			For i = 1 to .RecordCount
				If (Trim(.Fields("catName").Value) <> "No Category") Then
					'Create category Node
					Set xmlCategoryDetail = xmlDoc.createElement("category")
					xmlRoot.appendChild xmlCategoryDetail

					xmlCategoryDetail.setAttribute "ID", "root:" & .Fields("catID").Value
					Call addNode(xmlDoc, xmlCategoryDetail, "catParent", "root:" & 0)
					Call addNode(xmlDoc, xmlCategoryDetail, "catDepth", 0)
					Call addCDATA(xmlDoc, xmlCategoryDetail, "catName", .Fields("catName").Value)
					If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catDescription", .Fields("catDescription").Value)
					If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catURL", .Fields("catHttpAdd").Value)
					If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catImage", .Fields("catImage").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catIsActive", .Fields("catIsActive").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catHeirarchy", "")
					Call addNode(xmlDoc, xmlCategoryDetail, "catBottom", .Fields("catHasSubCategory").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catStatus", "")
					
				End If
				.MoveNext
			Next
			
			If .RecordCount = 0 Then
				'Create category Node
				Set xmlCategoryDetail = xmlDoc.createElement("category")
				xmlRoot.appendChild xmlCategoryDetail

				xmlCategoryDetail.setAttribute "ID", "new:0"
				Call addNode(xmlDoc, xmlCategoryDetail, "catParent", "root:" & 0)
				Call addNode(xmlDoc, xmlCategoryDetail, "catDepth", 0)
				Call addCDATA(xmlDoc, xmlCategoryDetail, "catName", "Enter you're first category")
				If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catDescription", "")
				If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catURL", "")
				If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catImage", "")
				Call addNode(xmlDoc, xmlCategoryDetail, "catIsActive", 0)
				Call addNode(xmlDoc, xmlCategoryDetail, "catHeirarchy", "")
				Call addNode(xmlDoc, xmlCategoryDetail, "catBottom", 0)
				Call addNode(xmlDoc, xmlCategoryDetail, "catStatus", "Insert")

			End If
			.Close
		End With
		Set prsAEProductCategories = Nothing
			
		'inner join used because SF doesn't always clean up after itself
		sql = "SELECT sfSub_Categories.subcatID, sfSub_Categories.subcatCategoryId, sfSub_Categories.Depth, sfSub_Categories.subcatName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatDescription, sfSub_Categories.subcatHttpAdd, sfSub_Categories.subcatImage, sfSub_Categories.bottom, sfSub_Categories.subcatIsActive" _
			& " FROM sfSub_Categories INNER JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
			& " ORDER BY sfSub_Categories.subcatCategoryId, sfSub_Categories.Depth, sfSub_Categories.subcatName;"

		Set prsAEProductCategories = GetRS(sql)
	'	Set prsAEProductCategories = Server.CreateObject("ADODB.RECORDSET")
		With prsAEProductCategories
			For i = 1 to .RecordCount
				If (Trim(.Fields("subcatName").Value) <> "No Category") Then	'Do not show No Category
					'Create category Node
					Set xmlCategoryDetail = xmlDoc.createElement("category")
					xmlRoot.appendChild xmlCategoryDetail

					xmlCategoryDetail.setAttribute "ID", "sub:" & .Fields("subcatID").Value
					If .Fields("Depth").Value = 0 Then
						Call addNode(xmlDoc, xmlCategoryDetail, "catParent", "root:" & .Fields("subcatCategoryId").Value)
					ElseIf .Fields("Depth").Value = 1 Then
						Call addNode(xmlDoc, xmlCategoryDetail, "catParent", "root:" & .Fields("subcatCategoryId").Value)
					Else
						pstrHeir = Trim(.Fields("CatHierarchy").Value)
						paHeir = Split(pstrHeir,"-")
						Call addNode(xmlDoc, xmlCategoryDetail, "catParent", "sub:" & paHeir(.Fields("Depth").Value - 2))
					End If

					Call addNode(xmlDoc, xmlCategoryDetail, "catDepth", .Fields("Depth").Value)
					Call addCDATA(xmlDoc, xmlCategoryDetail, "catName", .Fields("subcatName").Value)
					If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catDescription", .Fields("subcatDescription").Value)
					If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catURL", .Fields("subcatHttpAdd").Value)
					If cbytMode > 0 Then Call addCDATA(xmlDoc, xmlCategoryDetail, "catImage", .Fields("subcatImage").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catIsActive", .Fields("subcatIsActive").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catHeirarchy", .Fields("CatHierarchy").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catBottom", .Fields("bottom").Value)
					Call addNode(xmlDoc, xmlCategoryDetail, "catStatus", "")

				End If
				.MoveNext
			Next
			.Close
		End With
		Set prsAEProductCategories = Nothing
		
		AECategories_XML = xmlDoc.xml
		Set xmlDoc = Nothing
		Set xmlNode = Nothing
		Set xmlRoot = Nothing
		Set xmlCategoryDetail = Nothing

	End Function	'AECategories_XML

	'********************************************************************************************

	Function MakeXMLNode(strName,strValue)

	Dim pstrTemp
	Dim pstrTempValue

		'now to protect against special characters
		pstrTempValue = XMLEncode(strValue)
		
		pstrTemp = "  <" & strName & ">" & pstrTempValue & "</" & strName & ">" & vbcrlf
		MakeXMLNode = pstrTemp

	End Function	'MakeXMLNode

	'***************************************************************************************************************************************************************

	Function XMLEncode(byVal strText)

	Dim pstrOut

		pstrOut = Trim(strText & "")
		
		If False Then
			pstrOut = Replace(pstrOut, "&", "&#38;#38;")
			pstrOut = Replace(pstrOut, "<", "&#38;#60;")
			pstrOut = Replace(pstrOut, ">", "&#62;")

	'		pstrOut = Replace(pstrOut, Chr(34), "&#39;")
	'		pstrOut = Replace(pstrOut, "'", "&#34;")
		Else
			pstrOut = Replace(pstrOut, "&", "&amp;")
			pstrOut = Replace(pstrOut, "<", "&lt;")
			pstrOut = Replace(pstrOut, ">", "&gt;")

	'		pstrOut = Replace(pstrOut, Chr(34), "&quot;")
	'		pstrOut = Replace(pstrOut, "'", "&apos;")
		End If
		
		XMLEncode = pstrOut

	End Function	'XMLEncode

'********************************************************************************************
' Output Section
'********************************************************************************************

Sub OutputHTMLPage

'********************************************************************************************
'
'	USER CONFIGURATION
'
'********************************************************************************************

'this is where the top left corner of the category list view is
Const mlngLeft = 19
Const mlngTop = 203

'********************************************************************************************
'********************************************************************************************

Call WriteHeader("",True)
%>

<SCRIPT LANGUAGE="vbscript">
<!--
Option Explicit

Dim cbytMode
Dim cblnDebug

cbytMode = <%= cbytMode %>
cblnDebug = <%= cblnDebug %>

Dim mobjXMLDoc
Dim mstrcatID			'for uniqueness root: and sub: are added before the id
Dim mstrcatParent
Dim mbytDepth
Dim mstrcatName
Dim mstrcatDescription
Dim mstrcatURL
Dim mstrcatImage
Dim mblncatIsActive
Dim mstrCatHeir
Dim mstrcatBottom

Dim mlngNewCategoryCounter
Dim mstrStatus			'Insert, Update, Delete, or "" which means unchanged

Dim mlngX, mlngY

Dim mobjActiveXMLNode
Dim mobjActiveTreeNode
Dim mobjActiveTreeNodeForNew

'for drag/drop operation
Dim mblnDragging
Dim mblnValidDrop
Dim mobjSourceNode
Dim mobjTargetNode

'for edit operations
Dim mblnItemIsDirty
Dim mblnDataSetIsDirty


Const tvwChild = 4
Dim SourceNode

	'********************************************************************************************
	'********************************************************************************************

	Function CheckSaveChanges()

	Dim pblnResult
	Dim pblnResponse
	
		pblnResult = True
		
		If mblnItemIsDirty Then
			pblnResponse = msgbox("Do you wish to save your changes to " & mobjActiveTreeNode.Text & "?",vbYesNoCancel,"Save Changes?")
			Select Case pblnResponse
				Case vbYes	'Save Changes
					Call SaveChanges
				Case vbNo	'Abandon Changes
					Call AbandonChanges
				Case vbCancel	'return to original
					'this doesn't seem to work properly
					mobjActiveTreeNode.Selected = True
					pblnResult = False
			End Select
		End If
		
		CheckSaveChanges = pblnResult
		
	End Function	'CheckSaveChanges

	'********************************************************************************************

	Sub DeleteNodes(objNode,strKeys)
	
	Dim i
	Dim j
	Dim pobjChild
	
		If objNode.Children = 0 Then
			strKeys = strKeys & "|" & objNode.Key
		Else
			For i = objNode.Children To 1 Step -1
				Set pobjChild = objNode.Child
				For j = 1 To i-1
					Set pobjChild = pobjChild.Next
				Next 'j
				Call DeleteNodes(pobjChild,strKeys)
			Next 'i
			strKeys = strKeys & "|" & objNode.Key
		End If
	
	End Sub	'DeleteNodes

	'********************************************************************************************

	Sub DeleteCategory()

	Dim pblnResponse
	Dim pstrParent
	Dim pstrKeys
	Dim i,j
	Dim paryKeys
	Dim pobjNode
	Dim pobjnodeList

		pblnResponse = msgbox("Do you wish to delete category " & mobjActiveTreeNode.Text & "?",vbYesNoCancel,"Delete Category?")
		Select Case pblnResponse
			Case vbYes	'Delete

				Call DeleteNodes(mobjActiveTreeNode,pstrKeys)
				paryKeys = Split(pstrKeys,"|")

				For i = 1 To Ubound(paryKeys)
					'this takes care of the nodes on the tree
					Set pobjNode = GetXMLNodeByKey(paryKeys(i))
					Call LetElementValue(pobjNode, "catStatus", "Delete")
					Set pobjNode = Nothing
					
					'now for the nodes not added (ie the duplicate sub-categories)
					If Instr(1,paryKeys(i),"root:") <> 0 Then
						Set pobjnodeList = mobjXMLDoc.getElementsByTagName("category")
						For j = 0 To pobjnodeList.length - 1
							pstrParent = GetElementValue(pobjnodeList.Item(j), "catParent")
							If GetElementValue(pobjnodeList.Item(j), "catParent") = paryKeys(i) Then
								Call LetElementValue(pobjnodeList.Item(j), "catStatus","Delete")
								Set pobjnodeList = Nothing
								Exit For
							End If							
						Next 'j
						Set pobjnodeList = Nothing
					End If
						
				Next 'i
				
				Call MakeDataSetDirty(True)
				Call LoadListView
				Call SetFirstListItem
				
			Case vbNo, vbCancel	'Abandon Changes
				Call AbandonChanges
		End Select
		
	End Sub	'DeleteCategory

	'********************************************************************************************

	Sub CopyCategory()

	Dim pstrKeys
	Dim pobjNewNode
	Dim i,j
	Dim paryKeys
	Dim pobjNode
	Dim pobjnodeList
	Dim pobjDic

		Call DeleteNodes(mobjActiveTreeNode,pstrKeys)
		paryKeys = Split(pstrKeys,"|")

		'First Item keeps the parent
		i = 1
		Set pobjDic = CreateObject("Scripting.Dictionary")

		For i = 1 To Ubound(paryKeys)
			mlngNewCategoryCounter = mlngNewCategoryCounter + 1
			If pobjDic.Exists(paryKeys(i)) Then
				pobjDic.Item(paryKeys(i)) = "new:" & CStr(mlngNewCategoryCounter) & ":"
			Else
				pobjDic.Add paryKeys(i),"new:" & CStr(mlngNewCategoryCounter) & ":"
			End If
		Next 'i
		
		For i = 1 To Ubound(paryKeys)-1
			Set pobjNode = GetXMLNodeByKey(paryKeys(i))
			Set pobjNewNode = pobjNode.cloneNode(true)
				
			Call GetNodeValues(pobjNewNode)
			mstrcatID = pobjDic.Item(paryKeys(i))
			mstrcatParent = pobjDic.Item(mstrcatParent)
			mstrStatus = "Insert"
			
			pobjNewNode.attributes.item(0).nodeValue = mstrcatID
			Call LetNodeValues(pobjNewNode)
			mobjXMLDoc.documentElement.appendChild(pobjNewNode)
		
		Next 'i
				
		Set pobjNode = GetXMLNodeByKey(paryKeys(i))
		Set pobjNewNode = pobjNode.cloneNode(true)
		mobjXMLDoc.documentElement.appendChild(pobjNewNode)

		Call GetNodeValues(pobjNewNode)
		mstrcatID = pobjDic.Item(paryKeys(i))
		mstrcatName = "New Node"
		mstrStatus = "Insert"
		pobjNewNode.attributes.item(0).nodeValue = mstrcatID
		Call LetNodeValues(pobjNewNode)
		
		Call MakeDataSetDirty(True)
		Call LoadListView
		Call SetFirstListItem
		
		Set pobjDic = Nothing

	End Sub	'CopyCategory

	'********************************************************************************************

	Sub AddCategory()
	
	Dim pobjNewNode
	
		If Not CheckSaveChanges Then Exit Sub

		Set pobjNewNode = mobjActiveXMLNode.cloneNode(true)
		mobjXMLDoc.documentElement.appendChild(pobjNewNode)
		
		mlngNewCategoryCounter = mlngNewCategoryCounter + 1
		mstrcatID = "new:" & CStr(mlngNewCategoryCounter) & ":"
		mstrcatParent = mobjActiveTreeNodeForNew.Key
		mbytDepth = mbytDepth + 1
		mstrcatName = "New Node"
		mstrcatDescription = ""
		mstrcatURL = ""
		mstrcatImage = ""
		mblncatIsActive = True
		mstrStatus = "Insert"

		pobjNewNode.attributes.item(0).nodeValue = mstrcatID
		Set mobjActiveXMLNode = pobjNewNode

		Call LetNodeValues(pobjNewNode)

		TreeView1.Nodes.Add mstrcatParent, tvwChild, mstrcatID, mstrcatName
		
		Set mobjActiveTreeNode = TreeView1.Nodes.Item(TreeView1.Nodes.Count)
		Set mobjActiveTreeNodeForNew = mobjActiveTreeNode
		mobjActiveTreeNode.Selected = True
		Call SetFormValues
		
		Call MakeDataSetDirty(True)
		frmData.catName.focus
		frmData.catName.select
		
	End Sub	'AddCategory

	'********************************************************************************************

	Sub AutoUpdate()
	
		If document.frmData.chkAutoUpdate.checked Then
			msgbox "This will automatically save changes to the category as you make them." & vbcrlf _
				 & "You will lose the ability to use the reset button.",vbOKOnly,"Caution"
			frmData.btnReset.disabled = document.frmData.chkAutoUpdate.checked
			Call SaveChanges
		End If
	
	End Sub	'AutoUpdate

	'********************************************************************************************

	Sub ChangeItem()

		If document.frmData.chkAutoUpdate.checked Then
			SaveChanges()
		Else
			Call MakeItemDirty(True)
		End If
		Call MakeDataSetDirty(True)

	End Sub	'ChangeItem

	'********************************************************************************************

	Sub MakeItemDirty(blnDirty)
	
		If mblnItemIsDirty and Not blnDirty Then
			'clean item
			Call SetFormValues	'inserted for debugging to see status
			mblnItemIsDirty = False
			frmData.btnReset.disabled = (Not mblnItemIsDirty) OR document.frmData.chkAutoUpdate.checked
			frmData.btnUpdateItem.disabled = Not mblnItemIsDirty
		ElseIf Not mblnItemIsDirty and blnDirty Then
			'make item dirty
			
			mblnItemIsDirty = True
			frmData.btnReset.disabled = (Not mblnItemIsDirty) OR document.frmData.chkAutoUpdate.checked
			frmData.btnUpdateItem.disabled = Not mblnItemIsDirty
		End If

	End Sub

	'********************************************************************************************

	Sub MakeDataSetDirty(blnDirty)
	
		If mblnDataSetIsDirty and Not blnDirty Then
			'clean item
			
			mblnDataSetIsDirty = False
			frmData.btnUpdateDataset.disabled = Not mblnDataSetIsDirty
		ElseIf Not mblnDataSetIsDirty and blnDirty Then
			'make item dirty
			
			mblnDataSetIsDirty = True
			frmData.btnUpdateDataset.disabled = Not mblnDataSetIsDirty
		End If

	End Sub

	'********************************************************************************************

	Function ResetForm()
		Call SetFormValues
		Call MakeItemDirty(False)
	End Function	'ResetForm

	'********************************************************************************************

	Function SaveChanges()
		Call GetFormValues
		If Len(mstrStatus) = 0 Then mstrStatus = "Update"
		Call LetNodeValues(mobjActiveXMLNode)
		
		On Error Resume Next
		If IsObject(mobjActiveXMLNode) Then mobjActiveTreeNode.Text = mstrcatName
		Call MakeItemDirty(False)
	End Function	'SaveChanges

	'********************************************************************************************

	Function AbandonChanges()
		Call SetFormValues
		Call MakeItemDirty(False)
	End Function	'AbandonChanges

	'********************************************************************************************

Function ProcessUpdates()

'First populate the temporary listview

	Dim nodeList
	Dim altNodeList
	Dim i,j
	Dim pblnAllAdded
	Dim pblnTryAgain
	Dim plngCounter
	Dim pblnResponse

	'On Error Resume Next
	
		If Not CheckSaveChanges Then Exit Function

		SetStatusMessage "<h4>Processing data . . .</h4>"
		plngCounter = 0
		Set nodeList = mobjXMLDoc.getElementsByTagName("category")
		pblnAllAdded = False
		
		TreeView2.Nodes.Clear()
'		TreeView2.Nodes.Add ,,"root:0","Categories"

		Do While Not pblnAllAdded
			pblnTryAgain = False
		
			For i = 0 To nodeList.length - 1
				Call GetNodeValues(nodeList.Item(i))
				If mstrStatus <> "Delete" OR mstrStatus = "Delete" Then
					If mstrcatParent = "root:0" Then
						TreeView2.Nodes.Add , , mstrcatID, mstrcatID
					Else
						TreeView2.Nodes.Add mstrcatParent, tvwChild, mstrcatID, mstrcatID
					End If
					
					If Err.number <> 0 Then
						Select Case Err.number
							Case 35601	'Element Not Found
								pblnTryAgain = True
								'msgbox mstrcatID & ": Parent " & mstrcatParent
								Err.Clear
							Case 35602	'Key Is Not Unique in Collection
								'ignore this one
								Err.Clear
							Case Else
								MsgBox "Error " & Err.number & ": " & Err.Description,0,"Error"
								Err.Clear
						End Select
					End If
				Else
					'msgbox mstrcatName
				End If
			Next 'i

			pblnAllAdded = Not pblnTryAgain
			plngCounter = plngCounter + 1
			If plngCounter > (nodeList.length ^ 2) Then
				MsgBox "Bored"
				pblnAllAdded = True
			End If
		Loop
		
		'Expand all nodes
		For i = 1 To TreeView2.Nodes.Count
			TreeView2.Nodes.Item(i).Expanded = True
		Next 'i
	
		Set nodeList = Nothing

	'now start the analysis
	Dim pstrResult
	Dim pstrNode
	Dim paryNode
	Dim pobjXMLNode
	
	Dim pobjParentXMLNode
	Dim pobjTreeNode
	Dim pstrParentName
	Dim pstrCurrHier
	Dim pstrOrigDepth, pstrOrigHier, pstrOrigBottom
	Dim pstrCurrentStatus
	
	Dim pstrKey
	Dim pstrCatName
	Dim pstrParent
	Dim pstrHeir
	Dim pstrDepth
	Dim pstrBottom
	
	ReDim paryNodes(6,TreeView2.Nodes.Count)
	
	Const enCatNode_Key = 0
	Const enCatNode_Parent = 1
	Const enCatNode_Depth = 2
	Const enCatNode_CurrHier = 3
	Const enCatNode_Bottom = 4
	Const enCatNode_status = 5
	Const enCatNode_CatName = 6

	'0 - pstrKey
	'1 - pstrParent
	'2 - pstrDepth
	'3 - pstrCurrHier
	'4 - pstrBottom
	'5 - status
	'6 - pstrCatName
	
	SetStatusMessage "<h4>Starting analysis of changes . . .</h4>"

		pstrResult = pstrResult & PadItem("ID",11,True,False) _
								& PadItem("Name",15,True,False) _
								& PadItem("Parent",10,True,False) _
								& PadItem("Depth",7,True,False) _
								& PadItem("Heirarchy",25,True,False) _
								& PadItem("Bottom",8,True,False) _
								& PadItem("Status",8,True,False) _
								& vbcrlf

	For i = 1 to TreeView2.Nodes.Count
		Set pobjTreeNode = TreeView2.Nodes.Item(i)
		pstrNode = Replace(TreeView2.Nodes.Item(i).FullPath,"Categories-","")
		paryNode = Split(pstrNode,"-")
		
		'get the id
		pstrKey = TreeView2.Nodes.Item(i).Key
		paryNodes(enCatNode_Key,i) = pstrKey
		
		'get the name
		Set pobjXMLNode = GetXMLNodeByKey(pstrKey)
		pstrCatName = GetElementValue(pobjXMLNode, "catName")
		paryNodes(enCatNode_CatName,i) = pstrCatName
		
		'get the parent
		pstrParent = paryNode(0)
		paryNodes(enCatNode_Parent,i) = pstrParent
		'get the depth
		pstrDepth = UBound(paryNode)
		paryNodes(enCatNode_Depth,i) = pstrDepth
		
		'get the hierarchy
		paryNodes(enCatNode_CurrHier,i) = Replace(pstrNode,pstrParent & "-","")
		paryNodes(enCatNode_CurrHier,i) = Replace(paryNodes(enCatNode_CurrHier,i),pstrParent,"")

		'figure out hierarchy
		'msgbox pstrCatName & " - " & pstrKey & " - " & pstrBottom & " - " & TreeView2.Nodes.Item(i).Children
		If pstrDepth = 0 Then
			If Instr(1,pstrKey,"root:") = 1 Then
				For j = 1 to TreeView1.Nodes.Count
					If TreeView1.Nodes.Item(j).Key = pstrKey Then
						pstrBottom = Abs((TreeView1.Nodes.Item(j).Children = 0))
						Exit For
					End If
				Next 'j			
			ElseIf Instr(1,pstrKey,"new:") = 1 Then
				If (TreeView2.Nodes.Item(i).Children > 0) Then
					pstrBottom = 0
				Else
					pstrBottom = 1
				End If
'				pstrBottom = Abs((TreeView2.Nodes.Item(i).Children = 0))
				pstrCurrHier = ""
			Else
				pstrBottom = CInt(Abs((TreeView2.Nodes.Item(i).Children = 1)))
				paryNodes(enCatNode_Parent,i) = "root:0"	'added for dragging sub-category to root
				pstrCurrHier = ""
			End If
		ElseIf pstrDepth = 1 Then
			Set pobjParentXMLNode = GetXMLNodeByKey(paryNode(0))
			pstrParentName = GetElementValue(pobjParentXMLNode, "catName")
		
			'adjust depth down one if this is the sub-cat duplicate
			If pstrParentName = pstrCatName Then
				pstrDepth = pstrDepth - 1
				pstrCurrHier = Replace(pstrNode,pstrParent & "-","none-")

				For j = 1 to TreeView2.Nodes.Count
					If TreeView2.Nodes.Item(j).Key = paryNode(0) Then
						pstrBottom = Abs((TreeView2.Nodes.Item(j).Children = 1))
						Exit For
					End If
				Next 'j
			Else
				pstrCurrHier = Replace(pstrNode,pstrParent & "-","")
				pstrBottom = Abs((TreeView2.Nodes.Item(i).Children = 0))
			End If
			Set pobjParentXMLNode = Nothing
		Else
			pstrCurrHier = Replace(pstrNode,pstrParent & "-","")
			pstrBottom = Abs((TreeView2.Nodes.Item(i).Children = 0))
		End If
		
		paryNodes(enCatNode_Bottom,i) = pstrBottom
		
		pstrCurrHier = Replace(pstrCurrHier,"root:","")
		pstrCurrHier = Replace(pstrCurrHier,"sub:","")
		paryNodes(enCatNode_CurrHier,i) = pstrCurrHier
		
		
		'Now check for any changes missed by adding/moving nodes
		'deleting - changing nodes caught through normal methods
		'look at depth/hierarchy/bottom
		
		pstrOrigDepth = GetElementValue(pobjXMLNode, "catDepth")
		If Isnumeric(pstrOrigDepth) Then pstrOrigDepth = CInt(pstrOrigDepth)
		
		pstrOrigHier = GetElementValue(pobjXMLNode, "catHeirarchy")
		pstrOrigBottom = GetElementValue(pobjXMLNode, "catBottom")
		If Isnumeric(pstrOrigBottom) Then pstrOrigBottom = CInt(pstrOrigBottom)
		
		pstrCurrentStatus = GetElementValue(pobjXMLNode, "catStatus")
		
		If (pstrDepth <> pstrOrigDepth) Or (paryNodes(enCatNode_CurrHier,i) <> pstrOrigHier) Or (pstrBottom <> pstrOrigBottom) Then
		'msgbox pstrCatName & " (" & pstrKey & ") Depth - Original: " & pstrOrigDepth & " - New: " & pstrDepth
		'msgbox pstrCatName & " (" & pstrKey & ") Hier - Original: " & pstrOrigHier & " - New: " & paryNodes(enCatNode_CurrHier,i)
		'msgbox pstrCatName & " (" & pstrKey & ") Bottom - Original: " & pstrOrigBottom & " - New: " & pstrBottom

			'this line was commented out because it would erroneously update the none- subcategory
			'If Len(pstrCurrentStatus) = 0 Then pstrCurrentStatus = "Update"
		End If
		
		paryNodes(enCatNode_status,i) = pstrCurrentStatus
		Set pobjXMLNode = Nothing
		
		pstrResult = pstrResult & PadItem(pstrKey,11,True,False) _
								& PadItem(pstrCatName,15,True,False) _
								& PadItem(pstrParent,10,True,False) _
								& PadItem(pstrDepth,7,True,False) _
								& PadItem(paryNodes(enCatNode_CurrHier,i),25,True,False) _
								& PadItem(pstrBottom,8,True,False) _
								& PadItem(paryNodes(enCatNode_status,i),8,True,False) _
								& vbcrlf

'		pstrResult = pstrResult & TreeView2.Nodes.Item(i).Key & ": " & UBound(paryNode) & " - " & pstrNode & vbcrlf
'		msgbox TreeView2.Nodes.Item(i).FullPath,,TreeView2.Nodes.Item(i).Key
		If cblnDebug Then document.all("taData").value = pstrResult
	Next 'i

	For i = 1 to TreeView2.Nodes.Count
		If Len(paryNodes(enCatNode_status,i)) > 0 Then 
			Set pobjXMLNode = GetXMLNodeByKey(paryNodes(enCatNode_Key,i))
			LetElementValue pobjXMLNode,"catParent",paryNodes(enCatNode_Parent,i)
			LetElementValue pobjXMLNode,"catDepth",paryNodes(enCatNode_Depth,i)
			LetElementValue pobjXMLNode,"catHeirarchy",paryNodes(enCatNode_CurrHier,i)
			LetElementValue pobjXMLNode,"catBottom",paryNodes(enCatNode_Bottom,i)
			LetElementValue pobjXMLNode,"catStatus",paryNodes(enCatNode_status,i)
			Set pobjXMLNode = Nothing
		End If		
	Next
	If cbytMode = 2 Then document.all("divResult").innerHTML =  pstrResult
	Call UpdateDatabase
	
End Function	'ProcessUpdates

	'********************************************************************************************

	Function UpdateDatabase
		
	Dim pstrTemp
	
		If True Then
			SetStatusMessage "<h4>Sending changes to server . . .</h4>"
			pstrTemp = RetrieveRemoteData("sfCategoryAdminAE.asp?Action=SaveChanges",mobjXMLDoc.xml,True,False)
			If cbytMode = 2 Then document.all("divResult").innerHTML =  pstrTemp

			If Len(pstrTemp) = 0 Then
				SetStatusMessage "<h4>Changes saved to database . . .</h4>"
			Else
				SetStatusMessage "<h4><font color=red>" & pstrTemp & "</font></h4>"
			End If
			Call MakeDatasetDirty(False)
			
			Call LoadCategories

		Else
			document.frmXMLData.xmlData.value = mobjXMLDoc.xml
			If cblnDebug Then document.all("taXML").value = mobjXMLDoc.xml
			document.frmXMLData.submit
		End If

	End Function	'UpdateDatabase

	'********************************************************************************************

	Function PadItem(strSource,bytLen,blnPadRight,blnTrim)

		If Len(strSource) < bytLen Then
			If blnPadRight Then
				PadItem = strSource & Space(bytLen - Len(strSource))
			Else
				PadItem = Space(bytLen - Len(strSource)) & strSource
			End If
		ElseIf Len(strSource) > bytLen And blnTrim Then 
			PadItem = Left(strSource,bytLen)
		Else
			PadItem = strSource
		End If

	End Function	'PadItem

	'********************************************************************************************

	Function RetrieveRemoteData(strURL,strFormData,blnPostData,blnRandom)
	
	Dim pobjXMLHTTP
	
	'this is here to prevent data caching problem
	Dim pstrURL
	Dim pstrAppend
	
	If blnRandom Then
		Randomize   ' Initialize random-number generator.
		pstrAppend = Int(2147483647 * Rnd)
		If Instr(1,strURL,"?") > 0 Then
			pstrURL = strURL & "&RandomTrash=" & pstrAppend
		Else
			pstrURL = strURL & "?RandomTrash=" & pstrAppend
		End If
	Else
		pstrURL = strURL
	End If
	
	'set timeouts in milliseconds
	Const resolveTimeout = 1000
	Const connectTimeout = 1000
	Const sendTimeout = 1000
	Const receiveTimeout = 10000
	
	On Error Resume Next
	
		Set pobjXMLHTTP = CreateObject("Microsoft.XMLHTTP")
		With pobjXMLHTTP
			If blnPostData Then
				.setTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
				.open "POST", "", False
				.open "POST", pstrURL, False
				.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
				.send strFormData
			Else
				.open "GET", pstrURL, False
				.send
			End If
			.ContentType = "text/xml"
			RetrieveRemoteData  = .responseText
		end with
		set pobjXMLHTTP = nothing

	End Function	'RetrieveRemoteData

	'********************************************************************************************

	Function GetXMLNodeByKey(strKey)

	Dim nodeList
	Dim i
	
		Set nodeList = mobjXMLDoc.getElementsByTagName("category")

		For i =0 To nodeList.length - 1
			If nodeList.Item(i).attributes.item(0).nodeValue = strKey Then
				Set GetXMLNodeByKey = nodeList.Item(i)
				Exit Function
			End If
		Next 'i
		
		Set nodeList = Nothing
		
	End Function	'GetXMLNodeByKey

	'********************************************************************************************

	Function GetTreeNodeByKey(objTreeView, strKey)

	Dim i
	
		For i =1 To objTreeView.Nodes.Count
			If objTreeView.Nodes.Item(i).Key = strKey Then
				Set GetTreeNodeByKey = objTreeView.Nodes.Item(i)
				Exit Function
			End If
		Next 'i
		
	End Function	'GetTreeNodeByKey

	'********************************************************************************************

	Function GetElementValue(objNode, strKey)
	
	Dim i
	Dim e
	Dim pstrTemp
	
		For i = 0 To objNode.childNodes.length - 1
			Set e = objNode.childNodes.Item(i)
			If e.nodeName = strKey Then
				pstrTemp = e.text
				Exit For
			End If
		Next 'i
		GetElementValue = pstrTemp
		
	End Function	'GetElementValue

	'********************************************************************************************

	Sub GetNodeValues(objNode)
	
	Dim i
	Dim e
	
		mstrcatID = objNode.attributes.item(0).nodeValue
		For i = 0 To objNode.childNodes.length - 1
			Set e = objNode.childNodes.Item(i)
			Select Case e.nodeName
				Case "catID": mstrcatID = e.text
				Case "catParent": mstrcatParent = e.text
				Case "catDepth": mbytDepth = CInt(e.text)
				Case "catName": mstrcatName = e.text
				Case "catDescription": mstrcatDescription = e.text
				Case "catURL": mstrcatURL = e.text
				Case "catImage": mstrcatImage = e.text
				Case "catIsActive": mblncatIsActive = e.text
				Case "catStatus": mstrStatus = e.text
				Case "catHeirarchy": mstrCatHeir = e.text
				Case "catBottom": mstrcatBottom = e.text
			End Select
		Next 'i

	End Sub	'GetNodeValues

	'********************************************************************************************

	Sub LetNodeValues(objNode)
	
	Dim i
	Dim e
	
	If isObject(objNode) Then
		For i = 0 To objNode.childNodes.length - 1
			Set e = objNode.childNodes.Item(i)
			Select Case e.nodeName
				Case "catID": e.text = mstrcatID
				Case "catParent": e.text = mstrcatParent
				Case "catDepth": e.text = mbytDepth
				Case "catName": e.text = mstrcatName
				Case "catDescription": e.text = mstrcatDescription
				Case "catURL": e.text = mstrcatURL
				Case "catImage": e.text = mstrcatImage
				Case "catIsActive": e.text = mblncatIsActive
				Case "catStatus": e.text = mstrStatus
				Case "catHeirarchy": e.text = mstrCatHeir
				Case "catBottom": e.text = mstrcatBottom
			End Select
		Next 'i
	End If

	End Sub	'LetNodeValues

	'********************************************************************************************

	Sub LetElementValue(objNode, strKey, strValue)
	
	Dim i
	Dim e
	
		For i = 0 To objNode.childNodes.length - 1
			Set e = objNode.childNodes.Item(i)
			If e.nodeName = strKey Then 
				e.text = strValue
				Exit For
			End If
		Next 'i

	End Sub	'LetElementValue

	'********************************************************************************************

	Sub SetFormValues
	
		With document.frmData
			.catName.value = mstrcatName
			
			If cbytMode > 0 Then
				.catID.value = mstrcatID
				.catImage.value = mstrcatImage
				.catIsActive.checked = mblncatIsActive
				.catDescription.value = mstrcatDescription
				.catURL.value = mstrcatURL
			End If
			
			If cbytMode = 2 Then
				.catHeirarchy.value = mstrCatHeir
				.catBottom.value = mstrcatBottom
				.catParent.value = mstrcatParent
				.catDepth.value = mbytDepth
				.status.value = mstrStatus
			End If
		End With
	
	End Sub	'SetFormValues

	'********************************************************************************************

	Sub GetFormValues
	
		With document.frmData
			mstrcatName = .catName.value

			If cbytMode > 0 Then
				mstrcatID = .catID.value
				mstrcatImage = .catImage.value
				mblncatIsActive = .catIsActive.checked
				mstrcatDescription = .catDescription.value
				mstrcatURL = .catURL.value
			End If
			
			If cbytMode = 2 Then
				mstrCatHeir = .catHeirarchy.value
				mstrcatBottom = .catBottom.value
			End If
		End With
	
	End Sub	'SetFormValues

	'********************************************************************************************

	Function LoadListView
	
	Dim pblnMaxDepth
	Dim nodeList
	Dim altNodeList
	Dim i
	Dim pblnAllAdded
	Dim pblnTryAgain
	Dim plngCounter
	Dim paryCategories
	Dim pstrResult
	Dim pblnSuccess

	On Error Resume Next
	
		pblnSuccess = True
		plngCounter = 0
		pblnMaxDepth = 0
		Set nodeList = mobjXMLDoc.getElementsByTagName("category")
		pblnAllAdded = False
		
		TreeView1.Nodes.Clear()
		TreeView1.Nodes.Add ,,"root:0","Categories"
		pblnAllAdded = (nodeList.length = 0)
		ReDim paryCategories(2,nodeList.length - 1)
		Do While Not pblnAllAdded
		pblnTryAgain = False
		
			For i = 0 To nodeList.length - 1
				Call GetNodeValues(nodeList.Item(i))
				paryCategories(0,i) = mstrcatID
				paryCategories(1,i) = mstrcatName
				If CInt(mbytDepth) > pblnMaxDepth Then pblnMaxDepth = CInt(mbytDepth)
				
				If mstrStatus <> "Delete" Then
					If Instr(1,mstrCatHeir,"none-") = 0 Then	'Do not show duplicate sub-categories
						TreeView1.Nodes.Add mstrcatParent, tvwChild, mstrcatID, mstrcatName
						If Err.number <> 0 Then
							Select Case Err.number
								Case 35601	'Element Not Found
									pblnTryAgain = True
									Err.Clear
								Case 35602	'Key Is Not Unique in Collection
									'ignore this one
									Err.Clear
								Case Else
									MsgBox "Error " & Err.number & ": " & Err.Description,vbOKOnly,"Error"
									Err.Clear
							End Select
						Else
							paryCategories(2,i) = True
						End If
					Else
						paryCategories(2,i) = True
					End If
				Else
					'msgbox mstrcatName
				End If
			Next 'i
			pblnAllAdded = Not pblnTryAgain
			plngCounter = plngCounter + 1
			If plngCounter > (nodeList.length^2) Then
'			If plngCounter > (nodeList.length * pblnMaxDepth) Then
				For i = 0 To nodeList.length - 1
					If Not paryCategories(2,i) Then
						Call LetElementValue(nodeList.Item(i),"catStatus","Delete")
						pstrResult = pstrResult & paryCategories(1,i) & " does not have a corresponding category. " & vbcrlf
					End If
				Next 'i
				If MsgBox("There are sub-categories without corresponding categories." & vbcrlf & vbcrlf _
								& pstrResult & vbcrlf & vbcrlf & "Do you want to delete them? (Required to continue)" _
								,vbYesNo,"Error") = 6 Then
						Dim pstrTemp
						SetStatusMessage "<h4>Deleting extra sub-categories . . .</h4>"

						document.frmXMLData.xmlData.value = mobjXMLDoc.xml
						document.frmXMLData.submit
						Exit Function
						
						'pstrTemp = RetrieveRemoteData("CatchCategoriesXML.asp",mobjXMLDoc.xml,True)
						'MsgBox "Extra sub-categories have been deleted."
						'Call window_onload
				Else
					SetStatusMessage "<h3><font color=red>There are sub-categories without corresponding categories.<br>This script cannot continue.</font></h3>"
'					TreeView1.Nodes.Clear()
					pblnSuccess = False
				End If
				pblnAllAdded = True
			End If
		Loop
		
		'Expand all nodes
		For i = 1 To TreeView1.Nodes.Count
			TreeView1.Nodes.Item(i).Expanded = True
		Next 'i
	
		Set nodeList = Nothing
		LoadListView = pblnSuccess
		
	End Function	'LoadListView

	'********************************************************************************************

	Sub SetFirstListItem()
	
		If TreeView1.Nodes.Count > 1 Then 
			TreeView1.Nodes.Item(1).Expanded = True
			Set mobjActiveTreeNode = TreeView1.Nodes.Item(2)
			Set mobjActiveTreeNodeForNew = mobjActiveTreeNode
			mobjActiveTreeNode.Selected = True
			Set mobjActiveXMLNode = GetXMLNodeByKey(mobjActiveTreeNode.Key)
			Call GetNodeValues(mobjActiveXMLNode)
			Call SetFormValues
			frmData.btnDeleteCategory.disabled = False
			frmData.btnCopyCategory.disabled = False
		End If

	End Sub	'SetFirstListItem

	'********************************************************************************************

	Function GetListItem(lngX, lngY)
		Set GetListItem = TreeView1.HitTest(lngX,lngY)
	End Function	'GetListItem

	'********************************************************************************************

	Function IsListItem(lngX, lngY)

	Dim SourceNode

		Set SourceNode = TreeView1.HitTest(lngX,lngY)
		If SourceNode is Nothing Then
			IsListItem = "0"
		Else
			IsListItem = SourceNode.Key
		End If
		Set SourceNode = Nothing
		
	End Function	'IsListItem

	'********************************************************************************************

	Function SetStatusMessage(strMessage)
		document.all("divMessage").innerHTML = strMessage
	End Function	'SetStatusMessage

	'********************************************************************************************

	Function LoadCategories

	Dim pstrRawData
	Dim pblnResult
	
		pblnResult = True
		SetStatusMessage "<h4>Loading Categories . . .</h4>"

		set mobjXMLDoc = CreateObject("MSXML2.DOMDocument")
		mobjXMLDoc.async = false
		pstrRawData = RetrieveRemoteData("sfCategoryAdminAE.asp?Action=GetData","",False,True)
		If mobjXMLDoc.loadXML(pstrRawData) Then
			If LoadListView Then
				Call SetFirstListItem
				SetStatusMessage ""
			Else
				pblnResult = False
				SetStatusMessage "<h4><font color=red>Error loading List</font></h4>"
			End If
		Else
			MsgBox "Unable to load category information - Fail: LoadCategories-loadXML",vbOKOnly,"Error"
			SetStatusMessage "<h4><font color=red>Unable to load category information - Fail: LoadCategories-loadXML</font></h4>"
			If cblnDebug Then document.all("taXML").value = pstrRawData
			If cblnDebug Then document.all("taData").value = "pstrRawData:" & vbcrlf & pstrRawData
			pblnResult = False
		End If
		
		LoadCategories = pblnResult
		
	End Function	'LoadCategories

	'********************************************************************************************

	Sub window_onLoad

		'initialize variables
		mblnDragging = False
		mblnItemIsDirty = False
		mblnDataSetIsDirty = False
		mlngNewCategoryCounter = 0
		
		Call LoadCategories

	End Sub	'window_onLoad

	'********************************************************************************************

	Sub document_onmousemove

	Dim plngX, plngY

		plngX = 15 * (window.event.screenX - document.all("TreeView1Left").value) - 7 '+ window.screenLeft
		plngY = 15 * (window.event.screenY - document.all("TreeView1Top").value) - 7 '+ window.screenTop

		mlngX = plngX
		mlngY = plngY

		
		'for testing
		If cblnDebug Then
			'document.all("mlngX").innerText = plngX
			'document.all("mlngY").innerText = plngY
	
			'document.all("xPos").innerText = window.event.screenX
			'document.all("yPos").innerText = window.event.screenY
			
			Dim pobjTestNode
			Set pobjTestNode = TreeView1.HitTest(plngX,plngY)
			If pobjTestNode is Nothing Then
			window.status = ""
			Else
			window.status = pobjTestNode.Text & " - " & pobjTestNode.Key
			End If
			Set pobjTestNode = Nothing
		End If
		'end testing

		If mblnDragging Then
			Set SourceNode = TreeView1.HitTest(plngX,plngY)
			
			If SourceNode is Nothing Then
			'window.status = "x: " & plngX & " y: " & plngY
				window.status = "Nothing"
				document.all("Target").innerText = " "
				document.all("DragStatus").innerText = "Nothing to drop on"
				mblnValidDrop = False
			ElseIf mobjSourceNode = SourceNode Then
				document.body.style.cursor = "move"
				document.all("Target").innerText = SourceNode.Text
				document.all("DragStatus").innerText = "Can't drop on self"
				mblnValidDrop = False
			Else
				If SourceNode.Text = "No Category" Then
					document.all("Target").innerText = " "
					document.all("DragStatus").innerText = "You cannot drag to 'No Category'"
					mblnValidDrop = False
				Else
					document.body.style.cursor = "crosshair"
					window.status = SourceNode.Key
					If cblnDebug Then
						document.all("Target").innerText = SourceNode.Text & " - " & SourceNode.Key
					Else
						document.all("Target").innerText = SourceNode.Text
					End If
					document.all("DragStatus").innerText = "Ready to drop"
					Set mobjTargetNode = SourceNode
					mblnValidDrop = True
				End If
			End If
		End If

	End Sub	'document_onmousemove

	'********************************************************************************************

	Function isDragDropEnabled
	
	Dim pblnDropEnabled
	
	On Error Resume Next
	
		pblnDropEnabled = False
		If isObject(document.all("chkDropEnabled")) Then pblnDropEnabled = document.all("chkDropEnabled").checked
		
		isDragDropEnabled = pblnDropEnabled
		
	End Function	'isDragDropEnabled

	'********************************************************************************************

	Sub document_onmousedown

	Dim pstrKey
	Dim pstrPrompt
	Dim pobjSourceNode

'window.xPos.toString
'msgbox "Client Top: " & window.xPos.clientLeft & " - Client Left: " & window.yPos.s
'msgbox "Client Top: " & TreeView1.Style.posTop & " - Client Left: " & TreeView1.Style.posLeft
'msgbox "Client Top: " & TreeView1.Style.top & " - Client Left: " & TreeView1.Style.left
'msgbox document.all("divTreeView1").offsetTop & ": " & document.all("divTreeView1").offsetLeft
'msgbox document.all("divTreeView1").clientTop & ": " & document.all("divTreeView1").clientTop
'msgbox document.TreeView1.clientTop & ": " & document.TreeView1.clientLeft
'msgbox document.TreeView1.clientTop & ": " & document.TreeView1.clientLeft
'msgbox document.all("divTreeView1").offsetTop & ": " & document.all("divTreeView1").offsetLeft
'msgbox document.frmData.catBottom.parentElement.offsetTop & ": " & document.frmData.catBottom.parentElement.offsetLeft
'msgbox document.frmData.catBottom.clientTop & ": " & document.frmData.catBottom.clientLeft
'msgbox document.frmData.catBottom.clientTop & ": " & document.frmData.catBottom.clientLeft

		If window.event.button = 1 Then
			If isDragDropEnabled Then
				Set pobjSourceNode = GetListItem(mlngX,mlngY)
				If Not pobjSourceNode is Nothing Then
					If pobjSourceNode.Text = "No Category" Then
						document.all("DragStatus").innerText = "You cannot drag 'No Category'"
						document.body.style.cursor = "hand"
					Else
						Set mobjSourceNode = pobjSourceNode
						If cblnDebug Then
							document.all("Source").innerText = pobjSourceNode.Text & " - " & pobjSourceNode.Key
						Else
							document.all("Source").innerText = pobjSourceNode.Text
						End If
						mblnDragging = True
						document.body.style.cursor = "hand"
					End If
				End If
			Else
				document.all("DragStatus").innerText = "Drag & Drop Disabled"
			End If
		ElseIf window.event.button = 2 Then
			If document.activeElement.id = "TreeView1" Then
				document.all("TreeView1Left").value = window.event.screenX
				document.all("TreeView1Top").value = window.event.screenY
				
				pstrPrompt = "The position of the treeview has been updated to: " & vbcrlf _
						   & vbcrlf _
						   & "mlngLeft: " & document.all("TreeView1Left").value & vbcrlf _
						   & "mlngTop: " & document.all("TreeView1Top").value & vbcrlf _
						   & vbcrlf _
						   & "You may want to update your configuration as required."
						   
				MsgBox pstrPrompt,vbOkOnly,"Note"
				
			End If
		End If

	'move
	'hand
	'auto
	End Sub	'document_onmousedown

	'********************************************************************************************

	Function MoveNodeUp(strSourceKey, strTargetKey, blnMoveUp)
	
	Dim pstrParent
	Dim pstrKeys
	Dim pstrSourceTreeNode
	Dim pblnDropBelow
	Dim pstrTempXML

		Set pstrSourceTreeNode = GetTreeNodeByKey(TreeView1, strSourceKey)
		pstrParent = pstrSourceTreeNode.Parent.Key
		Call DeleteNodes(pstrSourceTreeNode,pstrKeys)
		pblnDropBelow = (Instr(1,pstrKeys,"|" & strTargetKey & "|") > 0)
		
		'If Not pblnDropBelow Then Exit Function
		MoveNodeUp = pblnDropBelow
		Exit Function
		
		'check and see if we're dragging a node below itself
		If pstrSourceTreeNode.Children > 0 Then
			pstrTempXML = Replace(mobjXMLDoc.xml,"<catParent>" & strSourceKey & "</catParent>","<catParent>" & pstrParent & "</catParent>")
			mobjXMLDoc.LoadXML pstrTempXML
		End If

	End Function	'MoveNodeUp

	'********************************************************************************************

	Sub document_onmouseup

	Dim pobjSourceTreeNode
	
		If mblnValidDrop Then

			mstrcatParent = mobjTargetNode.Key

			If cblnDebug Then
				document.all("TargetParent").innerText = mstrcatParent
				document.all("SourceID").innerText = mstrcatParent
			End If

			'update the XML document
			
			If MoveNodeUp(mstrCatID,mstrcatParent,False) Then
				MsgBox "You cannot drag a category below itself."
				Exit Sub
			End If
			
			Call LetElementValue(mobjActiveXMLNode, "catParent", mstrcatParent)
			Call LetElementValue(mobjActiveXMLNode, "catStatus", "Update")

			'now update the list view with the revised XML document
			LoadListView
			mobjTargetNode.Expanded = True
			Call MakeDataSetDirty(True)

		End If
		
		mblnValidDrop = False
		document.body.style.cursor = "auto"
		Set mobjSourceNode = Nothing
		document.all("Source").innerText = " "
		document.all("Target").innerText = " "
		Call CheckDropEnabled
		mblnDragging = False
		
	End Sub	'document_onmouseup

	'********************************************************************************************

	Sub CheckDropEnabled

		If document.all("chkDropEnabled").checked Then
			document.all("DragStatus").innerHTML = "Drag & Drop Enabled"
		Else
			document.all("DragStatus").innerHTML = "Drag & Drop Disabled"
		End If
  
	End Sub	'CheckDropEnabled

	'********************************************************************************************

	Sub TreeView1_Click

	Dim pstrKey
	Dim pobjSourceNode
	Dim pblnResponse

	On Error Resume Next

		If Not CheckSaveChanges Then Exit Sub
		
'		Set pobjSourceNode = GetListItem(mlngX,mlngY)
		Set pobjSourceNode = TreeView1.SelectedItem
		If mblnItemIsDirty Then
			If mobjActiveTreeNode.Key <> pobjSourceNode.Key Then
			End If
		End If
			
		If Not pobjSourceNode is Nothing Then
			Set mobjActiveTreeNodeForNew = pobjSourceNode
			pstrKey = pobjSourceNode.Key
			If pstrKey = "root:0" Then Exit Sub
			Set mobjActiveXMLNode = GetXMLNodeByKey(pstrKey)
			If Err.number <> 0 Then
				msgbox "Error " & Err.number & ": " & Err.Description & " (" & pstrKey & ")"
				Err.Clear
			End If
			
			Call GetNodeValues(mobjActiveXMLNode)
			Call SetFormValues
			Set mobjActiveTreeNode = pobjSourceNode
		End If

	End Sub	'TreeView1_Click

	'********************************************************************************************

-->
</SCRIPT>
<script language="javascript">


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
var xyz = "\\";

	if (gblnSwitch)
	{
	gobjImage.src = pstrFilePath;
	pstrItem = gobjImage.name.replace("img","");
	pstrHREF = pstrFilePath.replace(pstrBasePath,"");
	eval("document.frmData." + pstrItem).value = pstrHREF.replace(xyz,"/");
	gblnSwitch = false;
	theFile.value = "";
	}
}

</script>
<XML id="xmldso" async="false">
<ROOT>
<CATEGORIES>
<%
Response.Write AECategories
%>
</CATEGORIES>
</ROOT>
</XML>

<P align=left>
<TABLE cellpadding=0 cellspacing=4 border=0 ID="Table1">
<TR>
  <TD valign=top align=left>
	<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"  ID="Object1" VIEWASTEXT>
	  <PARAM name="LPKPath" value="SSLibrary/treeview.lpk">
	</OBJECT>

	<OBJECT id=TreeView1 style="LEFT: 0px; WIDTH: 300px; TOP: 0px; HEIGHT: 400px" height=400 width=300 classid=clsid:0713E8A2-850A-101B-AFC0-4210102A8DA7 name=TreeView1 codebase="http://activex.microsoft.com/controls/vb6/COMCTL32.CAB" VIEWASTEXT>
	  <PARAM NAME="_ExtentX" VALUE="7594">
	  <PARAM NAME="_ExtentY" VALUE="3784">
	  <PARAM NAME="_Version" VALUE="327682">
	  <PARAM NAME="HideSelection" VALUE="0">
	  <PARAM NAME="Indentation" VALUE="0">
	  <PARAM NAME="LabelEdit" VALUE="1">
	  <PARAM NAME="LineStyle" VALUE="1">
	  <PARAM NAME="PathSeparator" VALUE="\">
	  <PARAM NAME="Sorted" VALUE="1">
	  <PARAM NAME="Style" VALUE="7">
	  <PARAM NAME="ImageList" VALUE="">
	  <PARAM NAME="BorderStyle" VALUE="0">
	  <PARAM NAME="Appearance" VALUE="1">
	  <PARAM NAME="MousePointer" VALUE="0">
	  <PARAM NAME="Enabled" VALUE="1">
	  <PARAM NAME="OLEDragMode" VALUE="0">
	  <PARAM NAME="OLEDropMode" VALUE="0">
	</OBJECT>
  </TD>
<TD valign=top align=left>
<FORM id=frmData name=frmData action="CatchCategoriesXML.asp" onsubmit="return false;">
<INPUT type="hidden" id="xcatID" name="xcatID" value="">
<INPUT type="hidden" id="xmlData" name="xmlData" value="">
<input type=hidden id=strBaseHRef name=strBaseHRef Value="<%= mstrBaseHRef %>">
<input type=hidden id=strBasePath name=strBasePath Value="<%= mstrBasePath %>">
<span id=spantempFile style="display:none">
<input type=file id=tempFile name=tempFile onchange="ProcessPath(this);" size="20">
</span>
<TABLE width=95% cellpadding=0 cellspacing=2 border=0 ID="Table2">
  <TR>
    <TD align=center colspan=2><DIV id="divMessage"></DIV></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catName">Name: </LABEL></TD>
    <TD align=left><INPUT id="catName" name="catName" value="" onchange="ChangeItem()" onkeyup="ChangeItem()" maxlength=50 size=50></TD>
  </TR>
<% If cbytMode > 0 Then %>
  <TR>
    <TD align=right><LABEL for="catID">ID: </LABEL></TD>
    <TD align=left><INPUT id="catID" name="catID" value="" disabled></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catDescription">Description: </LABEL></TD>
    <TD align=left><textarea name="catDescription" id="catDescription" onkeyup="ChangeItem()" rows="5" cols="40"></textarea><a HREF="javascript:doNothing()" onClick="ChangeItem(); return openACE(document.frmData.catDescription);" title="Edit this field with the HTML Editor"><img SRC="images/prop.bmp" BORDER=0></a>
    </TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catURL">URL: </LABEL></TD>
    <TD align=left><INPUT id="catURL" name="catURL" value="" onkeyup="ChangeItem()" maxlength=100 size=50></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catImage">Image: </LABEL></TD>
    <TD align=left><INPUT id="catImage" name="catImage" value="" onchange="ChangeItem()" onkeyup="ChangeItem()" maxlength=255 size=50>&nbsp;<img style="cursor:hand" name="imgcatImage" id="imgcatImage" border="0" 
					onmouseover="DisplayTitle(this);return false;" onmouseout"return ClearTitle();" src="catImage" 
					onclick="ChangeItem(); return SelectImage(this);" 
					title="Click to edit this image">
    </TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catIsActive">Active: </LABEL></TD>
    <TD align=left><INPUT type="checkbox" id="catIsActive" name="catIsActive" onclick="ChangeItem()"></TD>
  </TR>
<% End If %>
<% If cbytMode = 2 Then %>
  <TR>
    <TD align=right><LABEL for="catDepth">Depth: </LABEL></TD>
    <TD align=left><INPUT id="catDepth" name="catDepth" value="" disabled></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catParent">Parent: </LABEL></TD>
    <TD align=left><INPUT id="catParent" name="catParent" value="" disabled></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catHeirarchy">Heirarchy: </LABEL></TD>
    <TD align=left><INPUT id="catHeirarchy" name="catHeirarchy" value="" disabled></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="catBottom">Bottom: </LABEL></TD>
    <TD align=left><INPUT id="catBottom" name="catBottom" value="" disabled></TD>
  </TR>
  <TR>
    <TD align=right><LABEL for="status">Status: </LABEL></TD>
    <TD align=left><INPUT id="status" name="status" value="" disabled></TD>
  </TR>
<% End If 'cbytMode = 2 %>
  <TR>
    <TD colspan=2 align=center>
      <INPUT class="butn" type=button onclick="ResetForm()" value="Reset" id=btnReset name=btnReset disabled>&nbsp;
      <INPUT class="butn" type=button onclick="SaveChanges()" value="Save Changes" id=btnUpdateItem name=btnUpdateItem disabled>&nbsp;
      <INPUT type="checkbox" id="chkAutoUpdate" name="chkAutoUpdate" onclick="AutoUpdate()"><LABEL for="chkAutoUpdate">&nbsp;Autosave&nbsp;Changes</LABEL>
    </TD>
  </TR>
  <TR>
    <TD colspan=2 align=center>
      <INPUT class="butn" type=button onclick="AddCategory()" value="Add Category" id=btnAddCategory name=btnAddCategory>&nbsp;
      <INPUT class="butn" type=button onclick="DeleteCategory()" value="Delete Category" id=(" & mstrcatID & "): " & plngNewCatID name=btnDeleteCategory disabled>&nbsp;
      <INPUT class="butn" type=button onclick="CopyCategory()" value="Copy Category" id=btnCopyCategory name=btnCopyCategory disabled>&nbsp;
    </TD>
  </TR>
  <TR>
    <TD colspan=2 align=center>
      <INPUT class="butn" type=button onclick="OpenHelp('ssHelpFiles/WebStoreManager/help_CategoryAdminAE.htm')" value="?" id=btnHelp name=btnHelp>&nbsp;
      <INPUT class="butn" type=button onclick="ProcessUpdates()" value="Update Database" id=btnUpdateDataset name=btnUpdateDataset disabled>&nbsp;
    </TD>
  </TR>

</FORM>

<!--
-->
  <TR><TD colspan=2 align=center><HR/></TD></TR>
  <TR><TH colspan=2 align=center>Category Drag & Drop</TH></TR>
  <TR>
  <TD colspan=2>
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 ID="Table3">
		<TR>
		  <TD align=right><B>Source:</B> </TD>
		  <TD align=left><DIV ID=Source>&nbsp;</DIV></TD>
		</TR>
		<TR>
		  <TD align=right><B>Target:</B> </TD>
		  <TD align=left><DIV ID=Target>&nbsp;</DIV></TD>
		</TR>
		<TR>
		  <TD align=right><B>Drag Status:</B> </TD>
		  <TD align=left><DIV ID=DragStatus>Drag & Drop Disabled</DIV></TD>
		</TR>
	</TABLE>
  </TD></TR>
  <TR>
    <TD>&nbsp;</TD><TD><INPUT type="checkbox" id="chkDropEnabled" name="chkDropEnabled" onclick="CheckDropEnabled()"><LABEL id=lblchkDropEnabled for="chkDropEnabled">&nbsp;Enable Drag & Drop</LABEL></TD>
  </TR>
<% If cblnDebug Then %>
  <TR>
    <TH colspan=2 align=center><HR></TH>
  </TR>
  <TR>
    <TD align=left>Target Parent: </TD>
    <TD><DIV ID=TargetParent>&nbsp;</DIV></TD>
  </TR>
  <TR>
    <TD align=left>Source ID: </TD>
    <TD><DIV ID=SourceID>&nbsp;</DIV></TD>
  </TR>
  <TR>
    <TD align=right>Top: </TD>
    <TD align=left><INPUT name="TreeView1Top" id="TreeView1Top" value="<%= mlngTop %>" size=5>&nbsp;<DIV ID=yPos></DIV></TD>
  </TR>
  <TR>
    <TD align=right>Left: </TD>
    <TD align=left><INPUT name="TreeView1Left" id="TreeView1Left" value="<%= mlngLeft %>" size=5>&nbsp;<DIV ID=xPos></DIV></TD>
  </TR>
<% Else %>
  <INPUT type=hidden name="TreeView1Top" id="Hidden1" value="<%= mlngTop %>">
  <INPUT type=hidden name="TreeView1Left" id="Hidden2" value="<%= mlngLeft %>">
<% End If %>

</td>
</TR>
</TABLE>
</P>

<FORM id=frmXMLData name=frmXMLData action="sfCategoryAdminAE.asp" method="post">
<INPUT type=hidden id=Action name=Action value="UpdateData">
<INPUT type=hidden id="Hidden3" name=xmlData value="">
</FORM>

<% If cblnDebug Then %>
<H3>Full Debugging Enabled</H3>

<TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0 ID="Table4">
<TR><TD></TD><TD>X</TD><TD>Y</TD></TR>
<TR><TD>Adj</TD><TD><DIV ID=mlngX></DIV></TD><TD><DIV ID=mlngY></DIV></TD></TR>
</TABLE>

<fieldset id="fsData">
	<legend>Data</legend>
	<TEXTAREA id=taData rows=15 cols=120 name="taData"></TEXTAREA>
</fieldset>

<fieldset id="fsSourceXML">
	<legend>SourceXML</legend>
	<TEXTAREA id=taXML rows=15 cols=120 NAME="taXML"><% Response.Write AECategories %>
</TEXTAREA>
</fieldset>

<fieldset id="fsTargetXML">
	<legend>TargetXML</legend>
	<DIV id=divResult></DIV>
</fieldset>
<% End If %>

<DIV id=divTreeView2 style="display: <% If cblnDebug Then Response.Write "x" %>none;">
<fieldset id="fsResultTreeView">
	<legend>ResultTreeView</legend>
	<OBJECT name=TreeView2 id=TreeView2 style="LEFT: 0px; WIDTH: 400px; TOP: 0px; HEIGHT: 400px" height=400 width=400 classid=clsid:0713E8A2-850A-101B-AFC0-4210102A8DA7 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="7594">
	<PARAM NAME="_ExtentY" VALUE="3784">
	<PARAM NAME="_Version" VALUE="327682">
	<PARAM NAME="HideSelection" VALUE="0">
	<PARAM NAME="Indentation" VALUE="0">
	<PARAM NAME="LabelEdit" VALUE="1">
	<PARAM NAME="LineStyle" VALUE="1">
	<PARAM NAME="PathSeparator" VALUE="-">
	<PARAM NAME="Sorted" VALUE="1">
	<PARAM NAME="Style" VALUE="7">
	<PARAM NAME="ImageList" VALUE="">
	<PARAM NAME="BorderStyle" VALUE="0">
	<PARAM NAME="Appearance" VALUE="1">
	<PARAM NAME="MousePointer" VALUE="1">
	<PARAM NAME="Enabled" VALUE="1">
	<PARAM NAME="OLEDragMode" VALUE="0">
	<PARAM NAME="OLEDropMode" VALUE="0">
	</OBJECT>
</fieldset>
</DIV>
<!--#include file="adminFooter.asp"-->
</BODY>
</HTML>
<% End Sub	'OutputHTMLPage %>
