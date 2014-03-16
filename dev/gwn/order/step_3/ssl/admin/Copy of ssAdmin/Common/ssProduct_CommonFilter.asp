<%
'********************************************************************************
'*   Product Manager Common Filter Version SF 5.0 		                        *
'*   Release Version:	1.07			                                        *
'*   Release Date:		October 19, 2003										*
'*   Revision Date:		June 7, 2004											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Const cblnShowLeadingProductID = True

Dim maryCustomTextFilter(3)

maryCustomTextFilter(0) = Array("Product ID", "sfProducts.prodID")
maryCustomTextFilter(1) = Array("Short Description", "sfProducts.prodShortDescription")
maryCustomTextFilter(2) = Array("Long Description", "sfProducts.prodDescription")
maryCustomTextFilter(3) = Array("Product Name", "sfProducts.prodName")

'******************************************************************************************************************************************************************

Dim mstrsqlWhere, mstrsqlHaving, mstrSortOrder,mstrOrderBy

Dim mlngManufacturerFilter,mlngCategoryFilter,mlngVendorFilter
Dim mbytCategoryFilter, mbytSubCategoryFilter
Dim mradTextSearch, mstrTextSearch, mcurUpperPrice, mcurLowerPrice, mradShowActive, mradShowOnsale, mradShowShipped
Dim mblnDetailInNewWindow
Dim mblnAutoShowDetailInWindow
Dim mblnShowUnassignedProducts
Dim mstrDateFieldToUse
Dim mstrStartDate
Dim mstrEndDate

	Dim mstrTemp
	mstrTemp = LoadRequestValue("chkDetailInNewWindow")

	If Instr(1,mstrTemp,",") > 0 Then
		mblnDetailInNewWindow = Trim(Right(mstrTemp,Len(mstrTemp)-Instr(1,mstrTemp,",")))
	Else
		mblnDetailInNewWindow = mstrTemp
	End If

	If Len(mblnDetailInNewWindow) = 0 Then 
		If mblnAutoShowDetailInWindow Then
			mblnDetailInNewWindow = 0
		Else
			mblnDetailInNewWindow = 1
		End If
	End If

'***********************************************************************************************

'MakeCombo_Saved(mrsCategory,"catName","catID",mclsProduct.prodCategoryId)
Function MakeCombo_Saved(strCombo, strSelected)

Dim pstrOutput
Dim pstrSessionCombo

	pstrSessionCombo = strCombo & "Combo"
	Select Case strCombo
		Case "category":
			If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select catID,catName from sfCategories Order By catName", "catName", "catID", "")
		Case "manufacturer":
			If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select mfgID,mfgName from sfManufacturers Order By mfgName", "mfgName", "mfgID", "")
		Case "vendor":
			If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select vendID,vendName from sfVendors Order By vendName", "vendName", "vendID", "")
		Case "product":
			If cblnShowLeadingProductID Then
				'If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID, prodID + ' - ' + prodName As prodNameID from sfProducts Order By prodName","prodNameID","prodID", "")
				If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID, prodID + ' - ' + prodName As prodNameID from sfProducts Order By prodID","prodNameID","prodID", "")
			Else
				'If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID, prodName from sfProducts Order By prodName","prodName","prodID", "")
				If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID, prodName from sfProducts Order By prodID","prodName","prodID", "")
			End If
			'If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID, prodName from sfProducts Order By prodName","prodName","prodID", "")
			If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID, prodName from sfProducts Order By prodID","prodName","prodID", "")
		Case "productName":
			If cblnShowLeadingProductID Then
				If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodID + ' - ' + prodName As prodNameID from sfProducts Order By prodName","prodNameID","", "")
			Else
				If Len(session(pstrSessionCombo)) = 0 Then session(pstrSessionCombo) = createCombo("Select prodName from sfProducts Order By prodName","prodName","", "")
			End If
	End Select
	
	If Len(strSelected) > 0 Then
		pstrOutput = Replace(session(pstrSessionCombo), "<option value=""" & strSelected & """>", "<option value=""" & strSelected & """ selected>")
	Else
		pstrOutput = session(pstrSessionCombo)
	End If

	MakeCombo_Saved = pstrOutput
	'Response.Write pstrOutput
	
End Function	'MakeCombo_Saved

'***********************************************************************************************

Sub resetCombo_Saved(byVal strCombo)

Dim pstrSessionCombo

	pstrSessionCombo = strCombo & "Combo"
	session.Contents.Remove(pstrSessionCombo)
	
End Sub	'resetCombo_Saved

'***********************************************************************************************

Function AECategoryFilter(ByVal bytCategoryFilter, ByVal bytsubCategoryFilter)

	If Len(Session("AECategoryFilter")) = 0 Then
		AECategoryFilter = Create_AECategoryFilter(bytCategoryFilter, bytsubCategoryFilter)
	Else
		AECategoryFilter = Session("AECategoryFilter")
	End If

End Function	'AECategoryFilter

'***********************************************************************************************

Function getCategoryName(byVal lngCategoryID, byVal strSeparator)

Dim i
Dim paryCategories
Dim prsAEProductCategories
Dim pstrOut
Dim pstrSQL

	If cblnSF5AE Then
		pstrSQL = "SELECT sfCategories.catName, sfSub_Categories.subcatName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatID" _
				& " FROM sfSub_Categories INNER JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
				& " WHERE subcatID=" & wrapSQLValue(lngCategoryID, False, enDatatype_number)
		Set prsAEProductCategories = GetRS(pstrSQL)
		If Not prsAEProductCategories.EOF Then
			paryCategories = Split(prsAEProductCategories.Fields("CatHierarchy").Value, "-")
			pstrOut = Trim(prsAEProductCategories.Fields("catName").Value & "")
			
			If paryCategories(i) <> "none" Then
				If UBound(paryCategories) > 0 Then
					For i = 0 To UBound(paryCategories) - 1
						pstrOut = pstrOut & strSeparator & getNameFromID("sfSub_Categories", "subcatID", "subcatName", False, paryCategories(i))
					Next 'i
				End If
				pstrOut = pstrOut & strSeparator & Trim(prsAEProductCategories.Fields("subcatName").Value & "")
			End If
		End If
		
		Call ReleaseObject(prsAEProductCategories)
	Else
		'Set prsAEProductCategories = GetRS(sql)
		pstrOut = getNameFromID("sfCategories", "catID", "catName", False, lngCategoryID)
	End If	'cblnSF5AE
	
	getCategoryName = pstrOut
	
End Function	'getCategoryName

'***********************************************************************************************

Function Create_AECategoryFilter(ByVal bytCategoryFilter, ByVal bytsubCategoryFilter)

Dim sql
Dim prsAEProductCategories
Dim i

	sql = "SELECT sfCategories.catID, sfSub_Categories.subcatID, sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName, sfSub_Categories.Depth, sfSub_Categories.bottom" _
		& " FROM sfSub_Categories RIGHT JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
		& " ORDER BY sfCategories.catName, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName"
'		& " ORDER BY sfCategories.catName, sfSub_Categories.Depth, sfSub_Categories.subcatName"
	'debugprint "sql",sql
	
	Set prsAEProductCategories = GetRS(sql)
	With prsAEProductCategories
		If Not .EOF Then

			'this creates the first dropdown
			Dim pstrTemp
			Dim plngCatID
			Dim pstrCatValue
			Dim pstrCatName
			Dim pbytCatDepth
			Dim pstrsubCatValue
			Dim pstrsubCatName
			Dim pstrSelected
			Dim pblnBottom
			
			Dim pstrOptionValue

			If Len(bytCategoryFilter) = 0 And Len(bytsubCategoryFilter) = 0 Then pstrSelected = "selected"
			pstrTemp = "<select id='CategoryFilter' name='CategoryFilter' size='10' multiple>" & vbcrlf _
					 & "  <option value='...' " & pstrSelected & ">- All -</Option>" & vbcrlf

			For i=1 to .RecordCount
				If len(.Fields("Depth").value & "") > 0 Then 'Then No sub-categories so don't display
					pbytCatDepth = .Fields("Depth").value
					
					If plngCatID <> Trim(.Fields("catID").value) Then
						plngCatID = Trim(.Fields("catID").value)
						
						pstrCatValue = .Fields("catID").value & "."
						pstrCatName = .Fields("catName").value
						pstrOptionValue = pstrCatValue & "." & pbytCatDepth & "." & .Fields("bottom").value
						
						pstrSelected = ""
						If pstrCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
						pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrOptionValue & chr(34) & pstrSelected & ">" & pstrCatName & "</option>" & vbcrlf 
						
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrOptionValue = pstrsubCatValue & "." & pbytCatDepth & "." & .Fields("bottom").value

							pstrSelected = ""
							If pstrsubCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrOptionValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					Else
						If pbytCatDepth > 0 Then
							pstrsubCatValue = .Fields("catID").value & "." & .Fields("subcatID").value
							pstrsubCatName = String(pbytCatDepth*2,"-") & ">" & " " & .Fields("subcatName").value
							pstrOptionValue = pstrsubCatValue & "." & pbytCatDepth & "." & .Fields("bottom").value
							
							pstrSelected = ""
							If pstrsubCatValue = CStr(bytCategoryFilter & "." & bytsubCategoryFilter) Then pstrSelected = "selected"
							pstrTemp = pstrTemp & "<option value=" & chr(34) & pstrOptionValue & chr(34) & pstrSelected & ">" & pstrsubCatName & "</option>" & vbcrlf
						End If
					End If
				End If
				
				.MoveNext
			Next
			pstrTemp = pstrTemp & "</select>" & vbcrlf
			
			
		Else
			pstrTemp = "<font color=red>No Categories</font><br>"
			pstrTemp = pstrTemp & "<select id=CatSource name=CatSource size=10 multiple>"
			pstrTemp = pstrTemp & "</select>"
		End If
	End With

	Create_AECategoryFilter = pstrTemp
	
End Function	'Create_AECategoryFilter

'******************************************************************************************************************************************************************

Function childCategories(byVal strCategoryIDs)

Dim i
Dim paryCategories
Dim pobjRS
Dim pstrSQL
Dim pstrTemp

	If Len(strCategoryIDs) > 0 Then
		paryCategories = Split(strCategoryIDs, ",")
		pstrSQL = paryCategories(0)
		For i = 1 To UBound(paryCategories)
			pstrSQL = pstrSQL & "," & paryCategories(i)
		Next 'i
		pstrSQL = "Select subcatID From sfSub_Categories Where CatHierarchy In (" & pstrSQL & ")"
		Set pobjRS = GetRS(pstrSQL)
		If pobjRS.EOF Then
			childCategories = strCategoryIDs
		Else
			pstrTemp = pobjRS.getString(, , , ",")
			'remove the trailing , which results
			pstrTemp = Left(pstrTemp, Len(pstrTemp) - 1)
			'childCategories = strCategoryIDs & "," & pstrTemp
			childCategories = strCategoryIDs & "," & childCategories(pstrTemp)
		End If
		Call ReleaseObject(pobjRS)
	End If

End Function	'childCategories

'******************************************************************************************************************************************************************

Function completeCategoryPath(byRef objXMLDoc, byVal categorySeparator)

Dim pobjXMLParent
Dim pstrCategoryName

	pstrCategoryName = objXMLDoc.childNodes(2).Text
	
	Set pobjXMLParent = objXMLDoc.parentNode
	If pobjXMLParent.nodeName = "Categories" Then
		completeCategoryPath = pstrCategoryName
	Else
		completeCategoryPath = completeCategoryPath(pobjXMLParent, categorySeparator) & categorySeparator & pstrCategoryName
	End If

End Function	'completeCategoryPath

'***********************************************************************************************

Function createCategoryXML(byRef objXMLDoc)

Dim en_CatFields_uid:	en_CatFields_uid = 0
Dim en_CatFields_ParentLevel:	en_CatFields_ParentLevel = 1
Dim en_CatFields_Name:	en_CatFields_Name = 2
Dim en_CatFields_ParentID:	en_CatFields_ParentID = 3
Dim en_CatFields_IsActive:	en_CatFields_IsActive = 4
Dim en_CatFields_Description:	en_CatFields_Description = 5
Dim en_CatFields_ImagePath:	en_CatFields_ImagePath = 6
Dim i						'As Integer
Dim pblnLocalDebug
Dim pblnParentFound
Dim plngIsActive			'As Integer
Dim plngParentID			'As Integer
Dim plngParentLevel			'As Integer
Dim plnguid					'As Integer
Dim pobjNodeList			'As String
Dim pobjRSCategories			'As String
Dim pobjXMLCategoriesNode	'As System.Xml.XmlElement
Dim pobjXMLCategoryNode		'As System.Xml.XmlElement
Dim pstrDescription			'As String
Dim pstrImagePath			'As String
Dim pstrName				'As String
Dim pstrSQL					'As String
Dim xmlNode

Exit Function
	pblnLocalDebug = False	'True	False
	If pblnLocalDebug Then Call resetCombo_Saved("category")
    pstrSQL = "SELECT uid, ParentLevel, Name, ParentID, IsActive, Description, ImagePath" _
            & " FROM Categories" _
            & " Where IsActive=1" _
            & " ORDER BY ParentLevel, Name"
    Set pobjRSCategories = Server.CreateObject("ADODB.RECORDSET")
    Set pobjRSCategories = GetRS(pstrSQL)

    With pobjRSCategories
		set objXMLDoc = server.CreateObject("MSXML2.DOMDocument.3.0")
		' Create processing instruction and document root
		Set xmlNode = objXMLDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'")
		Set xmlNode = objXMLDoc.insertBefore(xmlNode, objXMLDoc.childNodes.Item(0))
	   
		' Create document root
        Set pobjXMLCategoriesNode = objXMLDoc.CreateElement("Categories")
		'pobjXMLCategoriesNode.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
		Set objXMLDoc.documentElement = pobjXMLCategoriesNode

        Do While Not .EOF
            plnguid = Trim(.Fields(0).Value & "")
            plngParentLevel = Trim(.Fields(1).Value & "")
            pstrName = Trim(.Fields(2).Value & "")
            plngParentID = Trim(.Fields(3).Value & "")
            plngIsActive = Trim(.Fields(4).Value & "")
            pstrDescription = Trim(.Fields(5).Value & "")
            pstrImagePath = Trim(.Fields(6).Value & "")
            
            If pblnLocalDebug Then
				Response.Write "<fieldset><legend>Category</legend>"
				Response.Write "plnguid: " & plnguid & "<br />"
				Response.Write "pstrName: " & pstrName & "<br />"
				Response.Write "plngParentID: " & plngParentID & "<br />"
				Response.Write "</fieldset>"
            End If

            Set pobjXMLCategoryNode = CreateCategoryNode(plnguid, plngParentLevel, pstrName, plngParentID, plngIsActive, pstrDescription, pstrImagePath, objXMLDoc)
            If plngParentLevel = 0 Then
                pobjXMLCategoriesNode.AppendChild(pobjXMLCategoryNode)
            Else
				pblnParentFound = False
				Set pobjNodeList = objXMLDoc.GetElementsByTagName("Category")
				For i = 0 To pobjNodeList.Length - 1
					If pobjNodeList.Item(i).ChildNodes.Length > 0 Then
						If pobjNodeList.Item(i).ChildNodes.Item(en_CatFields_uid).text = plngParentID Then
							pobjNodeList.Item(i).AppendChild(pobjXMLCategoryNode)
							pblnParentFound = True
							Exit For
						End If
					End If
				Next 'i

				If Not pblnParentFound Then
					If pblnLocalDebug Then Response.Write "Unable to locate node " & plngParentID & "<br />"
				End If
            End If
            .MoveNext
        Loop
        .Close()

		If pblnLocalDebug Then
			Response.Write "<fieldset><legend>Category XML</legend><textarea rows=80 cols=120>" & vbcrlf & objXMLDoc.xml & vbcrlf & "</textarea></fieldset>"
		End If

		Set pobjXMLCategoryNode = Nothing
		Set pobjXMLCategoriesNode = Nothing
		Set pobjNodeList = Nothing
    End With    'pobjRSCategories
	Set pobjRSCategories = Nothing
	
	Set Session(cstrSession_CategoriesLocation) = objXMLDoc

    createCategoryXML = True

End Function    'createCategoryXML

'***********************************************************************************************

Function CreateCategoryNode(ByVal lnguid, ByVal lngParentLevel, ByVal strName, ByVal lngParentID, ByVal lngIsActive, ByVal strDescription, ByVal strImagePath, ByRef objXMLDoc)

    Dim pobjXMLCategoryNode				'As System.Xml.XmlElement
    Dim pobjXMLCategoryAttributeNode	'As System.Xml.XmlElement

    Set pobjXMLCategoryNode = objXMLDoc.CreateElement("Category")
    Call pobjXMLCategoryNode.SetAttribute("catID", lnguid)

    'add the uid
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("uid")
    pobjXMLCategoryAttributeNode.text = lnguid
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    'add the ParentLevel
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("ParentLevel")
    pobjXMLCategoryAttributeNode.text = lngParentLevel
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    'add the name
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("Name")
    pobjXMLCategoryAttributeNode.text = strName
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    'add the ParentID
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("ParentID")
    pobjXMLCategoryAttributeNode.text = lngParentID
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    'add the IsActive
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("IsActive")
    pobjXMLCategoryAttributeNode.text = lngIsActive
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    'add the Description
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("Description")
    pobjXMLCategoryAttributeNode.text = strDescription
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    'add the ImagePath
    Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("ImagePath")
    pobjXMLCategoryAttributeNode.text = strImagePath
    pobjXMLCategoryNode.AppendChild(pobjXMLCategoryAttributeNode)

    Set CreateCategoryNode = pobjXMLCategoryNode

	'	pstrTemp = "<font color=red>No Categories</font><br>"
	'	pstrTemp = pstrTemp & "<select id=CatSource name=CatSource size=10 multiple>"
	'	pstrTemp = pstrTemp & "</select>"

End Function    'CreateCategoryNode

'***********************************************************************************************

Function LoadTextFilter

Dim pstrsqlWhere

	'modified so could link in directly
	mradTextSearch = LoadRequestValue("radTextSearch")
	mstrTextSearch = trim(LoadRequestValue("TextSearch"))
	If (Len(mradTextSearch) > 0) And (Len(mstrTextSearch) > 0) Then
		pstrsqlWhere =  maryCustomTextFilter(mradTextSearch)(1) & " Like '%" & sqlSafe(mstrTextSearch) & "%'"
	End If

	LoadTextFilter = pstrsqlWhere
	
End Function    'LoadTextFilter

'******************************************************************************************************************************************************************

Sub LoadFilter

dim pstrsqlWhere

	'modified so could link in directly
	mlngManufacturerFilter = LoadRequestValue("ManufacturerFilter")
	mlngCategoryFilter = LoadRequestValue("CategoryFilter")
	mlngVendorFilter = LoadRequestValue("VendorFilter")
	mradTextSearch = LoadRequestValue("radTextSearch")
	mstrTextSearch = trim(LoadRequestValue("TextSearch"))
	mcurUpperPrice = trim(LoadRequestValue("UpperPrice"))
	mcurLowerPrice = trim(LoadRequestValue("LowerPrice"))
	mradShowActive = LoadRequestValue("radShowActive")
	mradShowOnsale = LoadRequestValue("radShowOnsale")
	mradShowShipped = LoadRequestValue("radShowShipped")
	mblnShowUnassignedProducts = CBool(LoadRequestValue("ShowUnassignedProducts") = "1")

	'Now check for subCategories
	Dim paryCat,paryTempCat
	Dim pstrTempCat, pstrTempSubCat
	Dim pbytTempDepth, pblnTempBottom

	paryTempCat = Split(mlngCategoryFilter,",")
	If Len(mlngCategoryFilter) > 0 Then
		'If isArray(paryTempCat) Then
		If InStr(1, mlngCategoryFilter, ",") > 0 OR InStr(1, mlngCategoryFilter, ".") > 0 Then
			If UBound(paryTempCat) >= 0 Then
				paryCat = Split(paryTempCat(0),".")
				pstrTempCat = Trim(paryCat(0))
				pstrTempSubCat = Trim(paryCat(1))
				pbytTempDepth = Trim(paryCat(2))
				pblnTempBottom = Trim(paryCat(3))
				For i = 1 To UBound(paryTempCat)
					paryCat = Split(paryTempCat(i),".")
					pstrTempCat = pstrTempCat & "," & Trim(paryCat(0))
					pstrTempSubCat = pstrTempSubCat & "," & Trim(paryCat(1))
					pbytTempDepth = pbytTempDepth & "," & Trim(paryCat(2))
					pblnTempBottom = pblnTempBottom & "," & Trim(paryCat(3))
				Next 'i
			
				mbytCategoryFilter = pstrTempCat
				mbytSubCategoryFilter = pstrTempSubCat
						
				'debugprint "mbytCategoryFilter",mbytCategoryFilter
				'debugprint "mbytSubCategoryFilter",mbytSubCategoryFilter
				'Response.Flush
			Else
				mbytCategoryFilter = mlngCategoryFilter
			End If
		Else
			mbytCategoryFilter = mlngCategoryFilter
		End If
	End If
	
	If Not isNumeric(mcurUpperPrice) Then mcurUpperPrice = ""
	If Not isNumeric(mcurLowerPrice) Then mcurLowerPrice = ""

	If True Then
		pstrsqlWhere = LoadTextFilter
	Else
	If len(mstrTextSearch) > 0 Then
		Select Case mradTextSearch
			Case "0"	'Do Not Include
			Case "1"	'ProductID
				pstrsqlWhere = "sfProducts.prodID Like '%" & sqlSafe(mstrTextSearch) & "%'"
			Case "2"	'Short Description
				pstrsqlWhere = "sfProducts.prodShortDescription Like '%" & sqlSafe(mstrTextSearch) & "%'"
			Case "3"	'Long Description
				pstrsqlWhere = "sfProducts.prodDescription Like '%" & sqlSafe(mstrTextSearch) & "%'"
			Case "4"	'Product Name
				pstrsqlWhere = "sfProducts.prodName Like '%" & sqlSafe(mstrTextSearch) & "%'"
		End Select	
	End If
	End If

	Select Case mradShowActive
		Case "0"	'Do Not Include
		Case "1"	'Active
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodEnabledIsActive=1"
			Else
				pstrsqlWhere = "prodEnabledIsActive=1"
			End If
		Case "2"	'Inactive
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodEnabledIsActive=0"
			Else
				pstrsqlWhere = "prodEnabledIsActive=0"
			End If
	End Select	
	
	Select Case mradShowShipped
		Case "0"	'Do Not Include
		Case "1"	'Shipped
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodShipIsActive=1"
			Else
				pstrsqlWhere = "prodShipIsActive=1"
			End If
		Case "2"	'Not Shipped
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodShipIsActive<>1"
			Else
				pstrsqlWhere = "prodShipIsActive<>1"
			End If
	End Select	
	
	Select Case mradShowOnsale
		Case "0"	'Do Not Include
		Case "1"	'Regular
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodSaleIsActive<>1"
			Else
				pstrsqlWhere = "prodSaleIsActive<>1"
			End If
		Case "2"	'Onsale
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and prodSaleIsActive=1"
			Else
				pstrsqlWhere = "prodSaleIsActive=1"
			End If
	End Select	
	
	If len(mlngManufacturerFilter) > 0 then
		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " and prodManufacturerId=" & mlngManufacturerFilter
		Else
			pstrsqlWhere = "prodManufacturerId=" & mlngManufacturerFilter
		End If
	End If


	If len(mbytCategoryFilter) > 0 OR mblnShowUnassignedProducts then

		If cblnSF5AE Then	' And False
			If mblnShowUnassignedProducts Then
				pstrTemp = "(sfSub_Categories.subcatCategoryId is Null)"
			Else
				Dim paryCategory, parysubCategory
				Dim paryDepth, paryBottom
				Dim i
				Dim pstrTemp
				Dim pstrCatSubCatChoice
				
				paryCategory = Split(mbytCategoryFilter,",")
				If Len(mbytsubCategoryFilter) = 0 Then
					parysubCategory = Array("")
				Else
					parysubCategory = Split(mbytsubCategoryFilter,",")
				End If

				If Len(pbytTempDepth) = 0 Then
					paryDepth = Array("")
				Else
					paryDepth = Split(pbytTempDepth,",")
				End If

				If Len(pblnTempBottom) = 0 Then
					paryBottom = Array("")
				Else
					paryBottom = Split(pblnTempBottom,",")
				End If

				'paryCategory = Split(mbytCategoryFilter,",")
				If isArray(paryCategory) Then
					For i = 0 To UBound(paryCategory)
						If Len(parysubCategory(i)) > 0 Then
							If paryBottom(i) = 1 Then
								pstrCatSubCatChoice = "(sfSub_Categories.subcatId=" & parysubCategory(i) & ")"
							Else
								If paryDepth(i) = 1 Then
									pstrCatSubCatChoice = "(sfSub_Categories.CatHierarchy Like '" & parysubCategory(i) & "-%')"
								Else
									pstrCatSubCatChoice = "(sfSub_Categories.CatHierarchy Like '%-" & parysubCategory(i) & "-%')"
								End If
							End If
							'Response.Write paryCategory(i) & "." & paryDepth(i) & " - " & paryBottom(i) & "<BR>"
						ElseIf Len(paryCategory(i)) > 0 Then
							pstrCatSubCatChoice = "(sfSub_Categories.subcatCategoryId=" & paryCategory(i) & ")"
						Else
							pstrCatSubCatChoice = ""
						End If
						
						If Len(pstrTemp) > 0 Then
							pstrTemp = pstrTemp & " OR " & pstrCatSubCatChoice
						Else
							pstrTemp = "( " & pstrCatSubCatChoice
						End If
					Next 'i
					If Request.Form("radCategoryInclude") <> "OR" Then mstrsqlHaving = " HAVING (Count(sfProducts.prodName)=" & CStr(UBound(paryCategory) + 1) & ")"
				Else
					pstrsqlWhere = "sfSub_Categories.subcatCategoryId=" & mbytCategoryFilter & ""
				End If
				pstrTemp = pstrTemp & ")"
			End If
		Else
			pstrTemp = "prodCategoryId=" & mbytCategoryFilter
		End If
		'debugprint "mstrsqlHaving",mstrsqlHaving
		'debugprint "pstrTemp",pstrTemp

		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " AND " & pstrTemp
		Else
			pstrsqlWhere = pstrTemp
		End If
	End If

	If len(mlngVendorFilter) > 0 then
		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " and prodVendorId=" & mlngVendorFilter
		Else
			pstrsqlWhere = "prodVendorId=" & mlngVendorFilter
		End If
	End If

	If len(mcurLowerPrice) > 0 then
		If cblnSQLDatabase Then
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and convert(money,prodPrice)>=" & mcurLowerPrice
			Else
				pstrsqlWhere = "convert(money,prodPrice)>=" & mcurLowerPrice
			End If
		Else
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and cCur(prodPrice)>=" & mcurLowerPrice
			Else
				pstrsqlWhere = "cCur(prodPrice)>=" & mcurLowerPrice
			End If
		End If
	End If

	If len(mcurUpperPrice) > 0 then
		If cblnSQLDatabase Then
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and convert(money,prodPrice)<=" & mcurUpperPrice
			Else
				pstrsqlWhere = "convert(money,prodPrice)<=" & mcurUpperPrice
			End If
		Else
			If len(pstrsqlWhere) > 0 Then
				pstrsqlWhere = pstrsqlWhere & " and cCur(prodPrice)<=" & mcurUpperPrice
			Else
				pstrsqlWhere = "cCur(prodPrice)<=" & mcurUpperPrice
			End If
		End If
	End If
	
	Dim pstrDateAddedField
	mstrDateFieldToUse = LoadRequestValue("DateFieldToUse")
	If Len(mstrDateFieldToUse) = 0 Then
		pstrDateAddedField = mstrDateFieldToUse
	Else
		pstrDateAddedField = "prodDateAdded"	'prodDateModified
	End If

	mstrStartDate = LoadRequestValue("StartDate")
	If len(mstrStartDate) > 0 then
		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " and (" & pstrDateAddedField & ">=" & wrapSQLValue(mstrStartDate & " 12:00:00 AM", False, enDatatype_date) & ")"
		Else
			pstrsqlWhere = "(" & pstrDateAddedField & ">=" & wrapSQLValue(mstrStartDate & " 12:00:00 AM", False, enDatatype_date) & ")"
		End If
	End If

	mstrEndDate = LoadRequestValue("EndDate")
	If len(mstrEndDate) > 0 then
		If len(pstrsqlWhere) > 0 Then
			pstrsqlWhere = pstrsqlWhere & " and (" & pstrDateAddedField & "<=" & wrapSQLValue(mstrEndDate & " 11:59:59 PM", False, enDatatype_date) & ")"
		Else
			pstrsqlWhere = "(" & pstrDateAddedField & "<=" & wrapSQLValue(mstrEndDate & " 11:59:59 PM", False, enDatatype_date) & ")"
		End If
	End If

	If len(pstrsqlWhere) > 0 then mstrsqlWhere = "where "  & pstrsqlWhere
	'debugprint "mstrsqlWhere",mstrsqlWhere
End Sub    'LoadFilter

'******************************************************************************************************************************************************************

Sub LoadSort

Dim pstrOrderBy

	mstrOrderBy = Request.Form("OrderBy")
	mstrSortOrder = Request.Form("SortOrder")

	If len(mstrOrderBy) = 0 Then mstrOrderBy = 1
	Select Case mstrOrderBy	'Order By
		Case "1"	'Product ID
			pstrOrderBy = "sfProducts.prodID"
		Case "2"	'Product Name
			pstrOrderBy = "sfProducts.prodName"
		Case "3"	'Price
			If cblnSQLDatabase Then
				pstrOrderBy = "convert(money,sfProducts.prodPrice)"
			Else
				pstrOrderBy = "cCur(sfProducts.prodPrice)"
			End If
		Case "4"	'Active
			pstrOrderBy = "sfProducts.prodEnabledIsActive"

		'this section added for updated Product Manager
		Case "prodID"	'Product ID
			pstrOrderBy = "sfProducts.prodID"
		Case "prodName"	'Product Name
			pstrOrderBy = "sfProducts.prodName"
		Case "prodPrice"	'Price
			If cblnSQLDatabase Then
				pstrOrderBy = "convert(money,sfProducts.prodPrice)"
			Else
				pstrOrderBy = "cCur(sfProducts.prodPrice)"
			End If
		Case "prodSalePrice"	'Price
			If cblnSQLDatabase Then
				pstrOrderBy = "convert(money,sfProducts.prodSalePrice)"
			Else
				pstrOrderBy = "cCur(sfProducts.prodSalePrice)"
			End If
		Case "prodDateAdded"	'Price
			pstrOrderBy = "prodDateAdded"
		Case "prodEnabledIsActive"	'Active
			pstrOrderBy = "sfProducts.prodEnabledIsActive"
	End Select	

	If len(pstrOrderBy) > 0 then mstrsqlWhere = mstrsqlWhere  & " Order By " & pstrOrderBy & " " & mstrSortOrder
	
End Sub    'LoadSort

'******************************************************************************************************************************************************************

Sub WriteProductFilter(byVal blnShowNewWindowChoice)

Dim i
%>
<script LANGUAGE=javascript>
<!--

function btnFilter_onclick(theButton)
{
var theForm = theButton.form;

  if (!isNumeric(theForm.LowerPrice,true,"Please enter a number for the price.")) {return(false);}
  if (!isNumeric(theForm.UpperPrice,true,"Please enter a number for the price.")) {return(false);}

  theForm.Action.value = "Filter";
  theForm.submit();
  return(true);
}

//-->
</script>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" rules="none" id="tblFilter">
  <tr>
    <th colspan=3 class="tblhdr">Find products which:</th>
  </tr>
  <tr>
    <td valign="top">
		<fieldset>
		<legend>Contain the text: </legend>
		  <div align=left>
        <input type="text" name="TextSearch" size="20" value="<%= EncodeString(mstrTextSearch,True) %>" ID="TextSearch" ondblclick="this.value='';">
        <p>in the field:</p>
        <% For i = 0 To UBound(maryCustomTextFilter) %>
		<input type="radio" value="<%= i %>" <% If mradTextSearch=CStr(i) Then Response.Write "checked" %> name="radTextSearch" ID="radTextSearch<%= i %>"><label for="radTextSearch<%= i %>"><%= maryCustomTextFilter(i)(0) %></label><br>
		<% Next 'i %>
        <input type="radio" value="" <% if mradTextSearch="" then Response.Write "Checked" %> name="radTextSearch" ID="radTextSearch"><label for="radTextSearch">Do Not Include</label>
        <br />
		  </div>
		</fieldset>
		<br />
		<fieldset>
		<legend>Have a status of: </legend>
		  <div align=left>
			<input type="radio" value="1" <% if mradShowActive="1" then Response.Write "Checked" %> name="radShowActive" id="radShowActive1"><label for="radShowActive1">Active</label>&nbsp;
			<input type="radio" value="2" <% if mradShowActive="2" then Response.Write "Checked" %> name="radShowActive" id="radShowActive2"><label for="radShowActive2">Inactive</label>&nbsp;
			<input type="radio" value="0" <% if (mradShowActive="0" or mradShowActive="") then Response.Write "Checked" %> name="radShowActive" id="radShowActive0"><label for="radShowActive0">All</label><br>

			<input type="radio" value="1" <% if mradShowOnsale="1" then Response.Write "Checked" %> name="radShowOnsale" id="radShowOnsale1"><label for="radShowOnsale1">Regular</label>&nbsp;
			<input type="radio" value="2" <% if mradShowOnsale="2" then Response.Write "Checked" %> name="radShowOnsale" id="radShowOnsale2"><label for="radShowOnsale2">On Sale</label>&nbsp;
			<input type="radio" value="0" <% if (mradShowOnsale="0" or mradShowOnsale="") then Response.Write "Checked" %> name="radShowOnsale" id="radShowOnsale0"><label for="radShowOnsale0">All</label><br>

			<input type="radio" value="1" <% if mradShowShipped="1" then Response.Write "Checked" %> name="radShowShipped" id="radShowShipped1"><label for="radShowShipped1">Shipped</label>&nbsp;
			<input type="radio" value="2" <% if mradShowShipped="2" then Response.Write "Checked" %> name="radShowShipped" id="radShowShipped2"><label for="radShowShipped2">Not Shipped</label>&nbsp;
			<input type="radio" value="0" <% if (mradShowShipped="0" or mradShowShipped="") then Response.Write "Checked" %> name="radShowShipped" id="radShowShipped0"><label for="radShowShipped0">All</label><br>
		  </div>
		</fieldset>
		<br />
		<fieldset>
		<legend>Are priced between: </legend>
		  <div align=left>
		    <table class="tbl" border=0 cellpadding=2 cellspacing=0>
		      <tr>
		        <td align=right><label for="LowerPrice">Min. Price:&nbsp;</label></td><td align=left><input type="text" id="LowerPrice" name="LowerPrice" size="10" value="<%= mcurLowerPrice %>" maxlength="15" ondblclick="this.value='';" onchange="if (!isNumeric(this,true,'Please enter a number')) {return(false);}"></td>
		      </tr>
		      <tr>
		        <td align=right><label for="UpperPrice">Max. Price:&nbsp;</label></td><td align=left><input type="text" id="UpperPrice" name="UpperPrice" size="10" value="<%= mcurUpperPrice %>" maxlength="15" ondblclick="this.value='';" onchange="if (!isNumeric(this,true,'Please enter a number')) {return(false);}"></td>
		      </tr>
		    </table>
		  </div>
		</fieldset>
		<br />
		<fieldset>
		<legend>Date Range between: </legend>
		  <div align=left>
		    <table class="tbl" border=0 cellpadding=2 cellspacing=0 id="Table1">
		      <tr>
		        <td align=right>&nbsp;</td>
		        <td align=left>
					<input type="radio" value="prodDateAdded" <% if mstrDateFieldToUse="prodDateAdded" or Len(mstrDateFieldToUse) = 0 Then Response.Write "Checked" %> name="DateFieldToUse" id="DateFieldToUse"><label for="DateFieldToUse">Added</label>&nbsp;
					<input type="radio" value="prodDateModified" <% if mstrDateFieldToUse="prodDateModified" then Response.Write "Checked" %> name="DateFieldToUse" id="DateFieldToUse1"><label for="DateFieldToUse1">Modified</label>&nbsp;
		        </td>
		      </tr>
		      <tr>
		        <td align=right><label for="StartDate">Start Date:&nbsp;</label></td>
		        <td align=left>
					<input id=StartDate name=StartDate value="<%= mstrStartDate %>" ondblclick="this.value='';">
					<a href="javascript:doNothing()" title="Select start date"
					onclick="setDateField(document.frmData.StartDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
					<img src="images/calendar.gif" border=0></a>
		        </td>
		      </tr>
		      <tr>
		        <td align=right><label for="EndDate">&nbsp;&nbsp;End Date:&nbsp;</label></td>
		        <td align=left>
					<input id=EndDate name=EndDate value="<%= mstrEndDate %>" ondblclick="this.value='';">
					<a href="javascript:doNothing()" title="Select end date"
					onclick="setDateField(document.frmData.EndDate); top.newWin = window.open('ssLibrary/calendar.html', 'cal', 'dependent=yes, width=210, height=230, screenX=200, screenY=300, titlebar=yes')">
					<img src="images/calendar.gif" border=0></a>
		        </td>
		      </tr>
		      <tr>
		        <td align=right></td><td align=left></td>
		        <td align=right></td><td align=left></td>
		      </tr>
		    </table>
		  </div>
		</fieldset>
	</td>
	<td valign="top" align="center">
	<p>Filter by Category<br>
<% If cblnSF5AE Then 
		Response.Write AECategoryFilter(mbytCategoryFilter, mbytsubCategoryFilter)
   Else %>
	<select size="1"  id="CategoryFilter" name="CategoryFilter">
	<% 	if len(mlngCategoryFilter) = 0 then
			Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
		else
			Response.Write "<option value=''>- All -</Option>" & vbcrlf
		end if
		'Call MakeCombo(mrsCategory,"catName","catID",mlngCategoryFilter)
		Response.Write MakeCombo_Saved("category", mlngCategoryFilter)
	%>
	</select>
<% End If %>

<% If cblnSF5AE Then %>
<br>
<input type="radio" value="OR" name="radCategoryInclude" id="radCategoryInclude1" <% If Request.Form("radCategoryInclude") <> "AND" Then Response.Write "checked" %>><label for="radCategoryInclude1">OR</label>&nbsp;
<input type="radio" value="AND" name="radCategoryInclude" id="radCategoryInclude2" <% If Request.Form("radCategoryInclude") = "AND" Then Response.Write "checked" %>><label for="radCategoryInclude2">AND</label>
<br><input type="checkbox" id="ShowUnassignedProducts" name="ShowUnassignedProducts" value="1" <% If mblnShowUnassignedProducts Then Response.Write "checked" %>><label for="ShowUnassignedProducts">Show only products not assigned to categories</label>
<% End If %>
    <p>Filter by Manufacturer<br>
	<select size="1"  id=ManufacturerFilter name=ManufacturerFilter>
<% 	if len(mlngManufacturerFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	'Call MakeCombo(mrsManufacturer,"mfgName","mfgID",mlngManufacturerFilter)
	Response.Write MakeCombo_Saved("manufacturer", mlngManufacturerFilter)
 %>
	</select></p>
	<p>Filter by Vendor<br>
	<select size="1"  id=VendorFilter name=VendorFilter>
<% 	if len(mlngVendorFilter) = 0 then
		Response.Write "<option value='' selected>- All -</Option>" & vbcrlf
	else
		Response.Write "<option value=''>- All -</Option>" & vbcrlf
	end if
	'Call MakeCombo(mrsVendor,"vendName","vendID",mlngVendorFilter)
	Response.Write MakeCombo_Saved("vendor", mlngVendorFilter)
%>
	</select>
	</td>
	<%
	Dim paryActions(23)
	
	paryActions(0) = Array("prodEnabledIsActive", "Active")
	paryActions(1) = Array("prodCountryTaxIsActive", "Apply Country Tax")
	paryActions(2) = Array("prodStateTaxIsActive", "Apply State Tax")
	paryActions(3) = Array("prodMessage", "Confirmation Message")
	paryActions(4) = Array("prodDateAdded", "Date Added")
	paryActions(5) = Array("prodDateModified", "Date Modified")
	paryActions(6) = Array("prodHeight", "Height")
	paryActions(7) = Array("prodImageLargePath", "Large Image")
	paryActions(8) = Array("prodLength", "Length")
	paryActions(9) = Array("prodDescription", "Long Description")
	paryActions(10) = Array("prodManufacturerId", "Manufacturer")
	paryActions(11) = Array("prodNamePlural", "Plural Name")
	paryActions(12) = Array("prodPrice", "Price")
	paryActions(13) = Array("prodName", "Product Name")
	paryActions(14) = Array("prodSalePrice", "Sale Price")
	paryActions(15) = Array("prodSaleIsActive", "Sale Is Active")
	paryActions(16) = Array("prodShipIsActive", "Shipped")
	paryActions(17) = Array("prodShip", "Shipping Cost")
	paryActions(18) = Array("prodShortDescription", "Short Description")
	paryActions(19) = Array("prodImageSmallPath", "Small Image")
	paryActions(20) = Array("prodWeight", "Weight")
	paryActions(21) = Array("prodWidth", "Width")
	paryActions(22) = Array("prodVendorId", "Vendor")
	paryActions(23) = Array("prodCategoryID", "Category")

	%>
	<td valign="top">
	  <input class="butn" id=btnFilter name=btnFilter type=button value="Apply Filter" onclick="btnFilter_onclick(this);"><br>
	  <% If blnShowNewWindowChoice Then %>
	  <input type="radio" name="chkDetailInNewWindow" id="chkDetailInNewWindow0" value="0" <%= isChecked(mblnDetailInNewWindow=0) %>>&nbsp;<label for="chkDetailInNewWindow0">Open detail in this window</label><br>
	  <input type="radio" name="chkDetailInNewWindow" id="chkDetailInNewWindow1" value="1" <%= isChecked(mblnDetailInNewWindow=1) %>>&nbsp;<label for="chkDetailInNewWindow1">Open detail in new window</label>
	  <hr />
	  <fieldset>
	    <legend>Update Selected Products</legend>
	  <select name="updateSelectedField" id="updateSelectedField">
	    <option value="">Select a field</option>
		<%
		For i = 0 To UBound(paryActions)
			Response.Write "<option value=""" & paryActions(i)(0) & """>" & paryActions(i)(1) & "</option>"
		Next 'i
		%>
	  </select>
	  <div>Set it to a value of</div>
	  <input type="text" name="updateSelectedValue" id="updateSelectedValue" value=""><br />
	  <input type=button class="butn" name="btnUpdateSelected" id="btnUpdateSelected" value="Update Selected" onclick="this.form.Action.value = 'UpdateSelected'; this.form.submit();"><br>
	  <hr />
	  <input class="butn" id="btnDeleteMarked" name=btnDeleteMarked type=button value="Delete Marked" onclick="this.form.Action.value = 'DeleteMarked'; this.form.submit();"><br>
	  <input class="butn" id="btnActivateMarked" name=btnActivateMarked type=button value="Activate Marked" onclick="this.form.Action.value = 'ActivateMarked'; this.form.submit();"><br>
	  <input class="butn" id="btnDeactivateMarked" name=btnDeactivateMarked type=button value="Deactivate Marked" onclick="this.form.Action.value = 'DeactivateMarked'; this.form.submit();"><br>
	  </fieldset>
	  <% End If %>
	</td>
  </tr>
</table>
<% End Sub	'WriteProductFilter %>