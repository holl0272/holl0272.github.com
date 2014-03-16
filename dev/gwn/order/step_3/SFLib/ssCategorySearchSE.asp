<%
'********************************************************************************
'*   sub-Category Search Tool for StoreFront 5.0 SE                             *
'*   Release Version   1.01.000                                                 *
'*   Release Date      July 5, 2002		                                        *
'*   Revision Date     October 30, 2003                                         *
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   Version 1.3                                                                *
'*    - Enhanced functionality to support only 1 level of categories            *
'*                                                                              *
'*   Version 1.2                                                                *
'*    - Bug fix: fixed some issues with Netscape 4.x		                    *
'*                                                                              *
'*   Version 1.1.1                                                              *
'*    - clean-up (starting row tag added to WriteSearchForm                     *
'*                                                                              *
'*   Version 1.1                                                                *
'*    - added Product Listing functionality                                     *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'***********************************************************************************************

Class ssCategorySE

Private prsCategory
Private pblnCategoriesLoaded

Private pstrCatOptions
Private pstrSubCatOptions
Private pstrSubSubCatOptions

Private pstrSubCatDBField
Private pstrSubSubCatDBField

Private pstrCatField
Private pstrSubCatField
Private pstrSubSubCatField

Private pstrSubCatText
Private pstrSubSubCatText

Private cstrLinkURL

Private cstrStartCat
Private cstrEndCat
Private cstrStartCatDetail
Private cstrEndCatDetail

Private cstrCatLinkStyle
Private cstrCatLinkStyle_Selected

Private cstrStartSubCat
Private cstrEndSubCat
Private cstrSubCatLIStyle
Private cstrSubCatLinkStyle
Private cstrSubCatLinkStyle_Selected

Private cstrStartSubSubCat
Private cstrEndSubSubCat
Private cstrSubSubCatLIStyle
Private cstrSubSubCatLinkStyle
Private cstrSubSubCatLinkStyle_Selected

Private cstrStartProduct
Private cstrEndProduct
Private cstrProductLIStyle
Private cstrProductLinkStyle

Private pstrFormName

Private cblnDispay
Private pstrCurrentCategory
Private pstrCurrentSubCategory
Private pstrCurrentSubSubCategory
Private pblnDisplayCurrentCategoryOnly
Private pblnHighlightSelected
Private pblnDisplayTopCategoriesOnly

Private pstrCurrentCategoryName
Private pstrCurrentCategoryImage
Private pstrCurrentCategoryDescription
Private pstrCurrentSubCategoryName
Private pstrCurrentSubSubCategoryName

Private pstrCacheName
Private plngCacheTime

Private pstrCategoryTrail
Private pstrCategoryTrailShowHome
Private pstrCategoryTrailShowHomeURL
Private pstrCurrentCategoryURL
Private pstrHRefClass
Private pstrHRefClass_SubCat
Private pstrHRefClass_SubSubCat
Private pstrStartFont
Private pstrEndFont
Private cstrTrailSpacerText

Private encatID
Private encatName
Private encatHttpAdd
Private enmfgID
Private enmfgName
Private enmfgHttpAdd
Private envendID
Private envendName
Private envendHttpAddr
Private paryCategories
Private pstrURLTemplate

'Start Sandshot Software Modification
'Purpose: Display Category heirarchy by Category/Manufacturer/Vendor

	'***********************************************************************************************

	Private Sub class_Terminate()
		On Error Resume Next
		If Not isEmpty(prsCategory) Then
			prsCategory.Close
			set prsCategory = nothing
		End If
	End Sub
	Private Sub class_Initialize()

		If cblnDebugCategorySearchTool Then
			Session("ssDebug_CategorySearchToolSE") = "True"
			Response.Write "<table border=0><tr><th colspan=2>Category Search Tool for StoreFront 5.0 SE</th></tr>"
			Response.Write "<tr><td>Release Version: </td><td>2.00.001</td></tr>"
			Response.Write "<tr><td>Release Date: </td><td>September 15, 2003</td></tr>"
			Response.Write "<tr><td colspan=""2""><hr></td></tr>"
			Response.Write "<tr><th colspan=""2"">Debugging enabled</th></tr>"
			Response.Write "</table>"
		End If

		pblnCategoriesLoaded = False
		pblnDisplayCurrentCategoryOnly = True
		pblnHighlightSelected = True
		cstrLinkURL = "search_results.asp?&txtsearchParamType=ALL"

		'for list style output
		cstrStartCat = "<ul class='clsMenu'>"
		cstrStartCatDetail = "<li class='clsMenu'>"
		cstrEndCatDetail = "</li>"
		cstrEndCat = "</ul>"
		
		'for tabular output
		'cstrStartCat = "<table cellpadding=0 cellspacing=0 border=1>"
		'cstrStartCatDetail = "<tr><td align=center>"
		'cstrEndCatDetail = "</td></tr>"
		'cstrEndCat = "</table>"
		
		cstrCatLinkStyle = "class='clsMenuNavigationCat'"
		cstrCatLinkStyle_Selected = "class='clsMenuNavigationCat_Selected'"
		
		cstrStartSubCat = "  <ul imagesrc='images/indent.gif'>"
		cstrStartSubCat = "  <ul class='clsMenu'>"
		cstrSubCatLIStyle = " onmouseover=" & Chr(34) & "this.style.listStyleType='circle'" & Chr(34) & " onmouseout=" & Chr(34) & "this.style.listStyleType=''" & Chr(34)
		cstrEndSubCat = "  </ul>"
		cstrSubCatLinkStyle = "class='clsMenuNavigationSubCat'"
		cstrSubCatLinkStyle_Selected = "class='clsMenuNavigationSubCat_Selected'"
		
		cstrStartSubSubCat = "    <ul>"
		cstrSubSubCatLIStyle = ""
		cstrEndSubSubCat = "    </ul>"
		cstrSubSubCatLinkStyle = "class='clsMenuNavigationSubCat'"
		cstrSubSubCatLinkStyle_Selected = "class='clsMenuNavigationSubCat_Selected'"
		
		cstrStartProduct = "      <ul>"
		cstrProductLIStyle = ""
		cstrEndProduct = "    </ul>"
		cstrProductLinkStyle = "class='clsMenuNavigationProduct'"
		
		pstrSubCatDBField = "mfg"
		pstrSubSubCatDBField = "vend"
		pstrSubSubCatDBField = ""
		pstrSubCatText = "-- All --"
		pstrSubSubCatText = "-- All --"
		pblnDisplayTopCategoriesOnly = False
		
		'Category Trail
		cstrTrailSpacerText = "&nbsp;>>&nbsp;"
		pstrCategoryTrailShowHome = "<span class=""categoryTrail"">Home</span>"
		pstrCategoryTrailShowHomeURL = "default.asp"

		pstrHRefClass = "clsMenuNavigationCat"
		pstrHRefClass_SubCat = "clsMenuNavigationCat"
		pstrHRefClass_SubSubCat = "clsMenuNavigationCat"
		pstrStartFont = "<span style=""vertical-align:middle"">" 'used to set font for category listing table
		pstrEndFont = "</span>"							'closing tag for category listing table font

		encatID = 0
		encatName = 1
		encatHttpAdd = 2
		enmfgID = 3
		enmfgName = 4
		enmfgHttpAdd = 5
		envendID = 6
		envendName = 7
		envendHttpAddr = 8

		plngCacheTime = 600	'in seconds
	End Sub

	Public Property Let FormName(strFormName)
		pstrFormName = strFormName
	End Property

	Public Property Let SubCatDBField(strSubCatDBField)
		pstrSubCatDBField = strSubCatDBField
	End Property

	Public Property Let SubSubCatDBField(strSubSubCatDBField)
		pstrSubSubCatDBField = strSubSubCatDBField
	End Property
	
	Public Property Let CurrentCategory(strCurrentCategory)
		pstrCurrentCategory = strCurrentCategory
	End Property
	
	Public Property Let CurrentSubCategory(strCurrentSubCategory)
		pstrCurrentSubCategory = strCurrentSubCategory
	End Property
	
	Public Property Let CurrentSubSubCategory(strCurrentSubSubCategory)
		pstrCurrentSubSubCategory = strCurrentSubSubCategory
	End Property
	
	Public Property Let DisplayCurrentCategoryOnly(blnDisplayCurrentCategoryOnly)
		pblnDisplayCurrentCategoryOnly = blnDisplayCurrentCategoryOnly
	End Property
	
	Public Property Get CategoryTrail
		CategoryTrail = pstrCategoryTrail
	End Property
	
	Public Property Let CategoryTrailShowHome(strCategoryTrailShowHome)
		pstrCategoryTrailShowHome = strCategoryTrailShowHome
	End Property
	
	Public Property Get CurrentCategoryName
		CurrentCategoryName = pstrCurrentCategoryName
	End Property
	
	Public Property Get CurrentCategoryImage
		CurrentCategoryImage = pstrCurrentCategoryImage
	End Property
	
	Public Property Get CurrentCategoryDescription
		CurrentCategoryDescription = pstrCurrentCategoryDescription
	End Property
	
	Public Property Let HighlightSelected(blnHighlightSelected)
		pblnHighlightSelected = blnHighlightSelected
	End Property
	
	Public Property Let CacheName(vntValue)
		pstrCacheName = vntValue
	End Property
	
	Public Property Let CacheTime(byVal vntValue)
		plngCacheTime = vntValue
	End Property
	
	Public Property Let DisplayTopCategoriesOnly(byVal vntValue)
		pblnDisplayTopCategoriesOnly = vntValue
	End Property
	
	Public Property Let URLTemplate(byVal vntValue)
		pstrURLTemplate = vntValue
	End Property

	'***********************************************************************************************
	
	Public Function createTrail()

	Dim pstrURL
	Dim pstrURL_Output
	Dim pstrCategoryName
	
		If LoadCategory(pstrCurrentCategory, pstrCurrentSubCategory, pstrCurrentSubSubCategory) Then
			'Add the home
			If Len(pstrCategoryTrailShowHome) > 0 Then pstrURL_Output = "<a href=" & Chr(34) & pstrCategoryTrailShowHomeURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & pstrCategoryTrailShowHome & pstrEndFont & "</a>"
			
			'Add the category
			If Len(CurrentCategoryName) > 0 Then
				pstrURL = "search_results.asp?txtsearchParamType=ALL&amp;txtsearchParamCat=" & pstrCurrentCategory & "&amp;txtsearchParamMan=ALL&amp;txtsearchParamVen=ALL"
				If Len(pstrURL_Output) > 0 Then
					pstrURL_Output = pstrURL_Output & cstrTrailSpacerText & "<a href=" & Chr(34) & pstrURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & StripHTML(CurrentCategoryName) & pstrEndFont & "</a>"
				Else
					pstrURL_Output = "<a href=" & Chr(34) & pstrURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & StripHTML(CurrentCategoryName) & pstrEndFont & "</a>"
				End If
			End If

			'Add the subcategory
			If Len(pstrCurrentSubCategoryName) > 0 Then
				pstrURL = "search_results.asp?txtsearchParamType=ALL" _
						& "&amp;txtsearchParamCat=" & pstrCurrentCategory _
						& "&amp;txtsearchParamMan=" & pstrCurrentSubCategory _
						& "&amp;txtsearchParamVen=ALL"
				pstrURL_Output = pstrURL_Output & cstrTrailSpacerText & "<a href=" & Chr(34) & pstrURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & pstrCurrentSubCategoryName & pstrEndFont & "</a>"
			End If
	
			'Add the subsubcategory
			If Len(pstrCurrentSubSubCategoryName) > 0 Then
				pstrURL = "search_results.asp?txtsearchParamType=ALL" _
						& "&amp;txtsearchParamCat=" & pstrCurrentCategory _
						& "&amp;txtsearchParamMan=" & pstrCurrentSubCategory _
						& "&amp;txtsearchParamVen=" & pstrCurrentSubSubCategory
				pstrURL_Output = pstrURL_Output & cstrTrailSpacerText & "<a href=" & Chr(34) & pstrCategoryTrailShowHomeURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & pstrCurrentSubSubCategoryName & pstrEndFont & "</a>"
			End If
		Else
			pstrURL_Output = ""
		End If
		
		createTrail = pstrURL_Output

	End Function	'createTrail

	'***********************************************************************************************

	Function LoadManufacturersByCategoryID(byVal lngCatID)

	Dim i
	Dim paryMfg
	Dim pobjCmd
	Dim pobjRS
	Dim pstrURL
	Dim pstrURL_Output

		If Len(lngCatID) = 0 Or Not isNumeric(lngCatID) Then
		
		Else
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "SELECT Distinct mfgID, mfgName" _
							& " FROM sfProducts INNER JOIN sfManufacturers ON sfProducts.prodManufacturerId = sfManufacturers.mfgID" _
							& " WHERE sfProducts.prodEnabledIsActive=1 And sfProducts.prodCategoryId=?" _
							& " Order By mfgName Asc"
				Set .ActiveConnection = cnn

				On Error Resume Next
				.Parameters.Append .CreateParameter("visitorID", adInteger, adParamInput, 4, lngCatID)
				Set pobjRS = .Execute
				If pobjRS.EOF Then
				
				Else
					paryMfg = pobjRS.GetRows()
					For i = 0 To UBound(paryMfg, 2)
						pstrURL = "search_results.asp?txtsearchParamType=ALL" _
								& "&amp;txtsearchParamCat=" & pstrCurrentCategory _
								& "&amp;txtsearchParamMan=" & paryMfg(0,i) _
								& "&amp;txtsearchParamVen=ALL"
						pstrURL_Output = "<a href=" & Chr(34) & pstrURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & paryMfg(1,i) & pstrEndFont & "</a>"
						Response.Write pstrURL_Output & "<br />"
					Next 'i
					'Response.Write "paryMfg: " & UBound()
				End If
				
				pobjRS.Close
				Set pobjRS = Nothing
			End With
			Set pobjCmd = Nothing
		End If

	End Function	'LoadManufacturersByCategoryID

	'***********************************************************************************************

	Function LoadCategoriesByManufacturerID(byVal lngMfgID)

	Const cblnLocalDebug = False

	Dim cstrURLTemplate
	Dim i
	Dim paryCategory
	Dim pclsSFSearch
	Dim pstrLink
	Dim pobjCmd
	Dim pobjRS
	
		If Len(pstrURLTemplate) = 0 Then
			cstrURLTemplate = "<a href=""search_results_manufacturer.asp?{QueryString}"" class=""clsMenuNavigationCat"" title=""{Title}"">{Text}</a>"
		Else
			cstrURLTemplate = pstrURLTemplate
		End If
	
		If cblnLocalDebug Then Response.Write "<fieldset><legend>LoadCategoriesByManufacturerID</legend>cstrURLTemplate: " & Server.HTMLEncode(cstrURLTemplate) & "<br />lngMfgID: " & lngMfgID & "<br />"

		If Len(lngMfgID) = 0 Or Not isNumeric(lngMfgID) Then
		
		Else
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "SELECT DISTINCT catID, catName, catDescription" _
							 & " FROM sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID" _
							 & " WHERE sfProducts.prodEnabledIsActive=1 And sfProducts.prodManufacturerId=?" _
							 & " Order By catName Asc"
				Set .ActiveConnection = cnn
				On Error Resume Next
				.Parameters.Append .CreateParameter("prodManufacturerId", adInteger, adParamInput, 4, lngMfgID)
				Set pobjRS = .Execute
				
				If cblnLocalDebug Then Response.Write "EOF: " & pobjRS.EOF & "<br />"
				If pobjRS.EOF Then
				
				Else
					paryCategory = pobjRS.GetRows()
					If cblnLocalDebug Then Response.Write "Count: " & UBound(paryCategory, 2) & "<br />"
					Set pclsSFSearch = New sfSearch
					For i = 0 To UBound(paryCategory, 2)
						If cblnLocalDebug Then Response.Write i & ": " & paryCategory(0, i) & "<br />"
						pclsSFSearch.txtsearchParamCat = paryCategory(0, i)
						pclsSFSearch.txtsearchParamMan = lngMfgID
						pstrLink = cstrURLTemplate
						pstrLink = Replace(pstrLink, "{QueryString}", pclssfSearch.SearchLinkParameters(False))
						pstrLink = Replace(pstrLink, "{Title}", "")
						pstrLink = Replace(pstrLink, "{Text}", paryCategory(1, i))

						Response.Write pstrLink & "<br />"
					Next 'i
					Set pclsSFSearch = Nothing
				End If	'pobjRS.EOF
				
				pobjRS.Close
				Set pobjRS = Nothing
			End With
			Set pobjCmd = Nothing
		End If
		If cblnLocalDebug Then Response.Write "</fieldset>"

	End Function	'LoadCategoriesByManufacturerID

	'***********************************************************************************************

	Public Function LoadCategory(byVal strCurrentCategory, byVal strCurrentSubCategory, byVal strCurrentSubSubCategory)
	
	Dim pobjRS
	Dim pstrSQL
	Dim pstrSQLOrderBy
	Dim pstrSQLWhere
	
		pstrCurrentCategory = strCurrentCategory
		pstrCurrentSubCategory = strCurrentSubCategory
		pstrCurrentSubSubCategory = strCurrentSubSubCategory
		
		If pstrCurrentCategory = "ALL" Then pstrCurrentCategory = 0
	
		If Len(pstrCurrentCategory) = 0 Then 
			LoadCategory = False
			Exit Function
		End If
		
		pstrSQLWhere = " WHERE " & makeSQLUpdate("catID", pstrCurrentCategory, False, enDatatype_number)
		pstrSQLOrderBy = " ORDER BY sfCategories.catName"
		
		If Len(pstrSubCatDBField) > 0 And strCurrentSubCategory <> "ALL" Then
			pstrSQLWhere = pstrSQLWhere & " And " & makeSQLUpdate("mfgID", pstrCurrentSubCategory, False, enDatatype_number)
			pstrSQLOrderBy = pstrSQLOrderBy & ", " & pstrSubCatDBField& "Name"
		End If
		
		If Len(pstrSubSubCatDBField) > 0 And strCurrentSubSubCategory <> "ALL" Then
			pstrSQLWhere = pstrSQLWhere & " And " & makeSQLUpdate("vendID", pstrCurrentSubSubCategory, False, enDatatype_number)
			pstrSQLOrderBy = pstrSQLOrderBy & ", " & pstrSubSubCatDBField& "Name"
		End If

		pstrSQL = "SELECT catID, catName, catDescription, catImage, mfgID, mfgName, mfgHttpAdd, vendID, vendName, vendHttpAddr" _
				& " FROM sfCategories INNER JOIN (sfVendors INNER JOIN (sfManufacturers INNER JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId) ON sfCategories.catID = sfProducts.prodCategoryId" _
				& pstrSQLWhere _
				& pstrSQLOrderBy
		Set pobjRS = GetRS(pstrSQL)

		If pobjRS.EOF Then
			LoadCategory = False
		Else
			pstrCurrentCategoryName = Trim(pobjRS.Fields("catName").Value)
			pstrCurrentCategoryImage = Trim(pobjRS.Fields("catImage").Value)
			pstrCurrentCategoryDescription = Trim(pobjRS.Fields("catDescription").Value)
			
			If Len(pstrSubCatDBField) > 0 And strCurrentSubCategory <> "ALL" Then pstrCurrentSubCategoryName = Trim(pobjRS.Fields("mfgName").Value)
			If Len(pstrSubSubCatDBField) > 0 And strCurrentSubSubCategory <> "ALL" Then pstrCurrentSubSubCategoryName = Trim(pobjRS.Fields("vendName").Value)
			
			If False Then
				Response.Write "<fieldset><legend>Sub LoadCategory</legend>"
				Response.Write "pstrCurrentCategoryName: " & pstrCurrentCategoryName & " (" & pstrCurrentCategory & ")" & "<br />"
				Response.Write "pstrCurrentSubCategoryName: " & pstrCurrentSubCategoryName & " (" & strCurrentSubCategory & ")" & "<br />"
				Response.Write "pstrCurrentSubSubCategoryName: " & pstrCurrentSubSubCategoryName & " (" & strCurrentSubSubCategory & ")" & "<br />"
				Response.Write "</fieldset>"
			End If

			LoadCategory = True
		End If
		Call closeObj(pobjRS)
		
	End Function	'LoadCategory

	'***********************************************************************************************

	Public Sub LoadCategories()
	
	Dim pstrSQL
	Dim pstrSQLOrderBy
	
		If pblnDisplayTopCategoriesOnly Then
			Call LoadCategories_Basic
			Exit Sub
		End If

		'Response.Write "<h4>Application caching of categories disabled for testing</h4>"
		'Call removeFromCache("ssCategorySearch" & pstrCacheName)
		If isCacheItemExpired("ssCategorySearch" & pstrCacheName) Then
			If cblnDebugCategorySearchTool Then Response.Write "<b>Loading Categories from database</b><br />"
			
			pstrSQLOrderBy = " ORDER BY sfCategories.catName"
			If Len(pstrSubCatDBField) > 0 Then pstrSQLOrderBy = pstrSQLOrderBy & ", " & pstrSubCatDBField& "Name"
			If Len(pstrSubSubCatDBField) > 0 Then pstrSQLOrderBy = pstrSQLOrderBy & ", " & pstrSubSubCatDBField& "Name"

			pstrSQL = "SELECT catID, catName, catHttpAdd, mfgID, mfgName, mfgHttpAdd, vendID, vendName, vendHttpAddr" _
					& " FROM sfCategories INNER JOIN (sfVendors INNER JOIN (sfManufacturers INNER JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId) ON sfCategories.catID = sfProducts.prodCategoryId" _
					& " WHERE sfProducts.prodEnabledIsActive=1" _
					& pstrSQLOrderBy
			'Response.Write "pstrSQL = " & pstrSQL & "<br />"
			Call DebugRecordTime("Loading categories . . .")
			Set prsCategory = CreateObject("ADODB.RECORDSET")
			With prsCategory
				.ActiveConnection = cnn
				.CursorLocation = 2 'adUseClient
				.CursorType = 3 'adOpenStatic
				.LockType = 1 'adLockReadOnly
				.Source = pstrSQL
				.Open
				'Response.Write "RecordCount = " & .RecordCount & "<br />"
				
				If Not .EOF Then
					paryCategories = .GetRows()
					
					Dim i
					Dim plngNumCategories
					
					plngNumCategories = UBound(paryCategories, 2)

					For i = 0 To plngNumCategories
						If Len(paryCategories(encatID, i) & "") > 0 And isNumeric(paryCategories(encatID, i) & "") Then
							paryCategories(encatID, i) = CLng(paryCategories(encatID, i))
						Else
							paryCategories(encatID, i) = 1
						End If
						If Len(paryCategories(enmfgID, i) & "") > 0 And isNumeric(paryCategories(enmfgID, i) & "") Then
							paryCategories(enmfgID, i) = CLng(paryCategories(enmfgID, i))
						Else
							paryCategories(enmfgID, i) = 1
						End If
						If Len(paryCategories(envendID, i) & "") > 0 And isNumeric(paryCategories(envendID, i) & "") Then
							paryCategories(envendID, i) = CLng(paryCategories(envendID, i))
						Else
							paryCategories(envendID, i) = 1
						End If
						
						paryCategories(encatName, i) = Trim(paryCategories(encatName, i) & "")
						paryCategories(encatHttpAdd, i) = Trim(paryCategories(encatHttpAdd, i) & "")
						paryCategories(enmfgName, i) = Trim(paryCategories(enmfgName, i) & "")
						paryCategories(enmfgHttpAdd, i) = Trim(paryCategories(enmfgHttpAdd, i) & "")
						paryCategories(envendName, i) = Trim(paryCategories(envendName, i) & "")
						paryCategories(envendHttpAddr, i) = Trim(paryCategories(envendHttpAddr, i) & "")
					
					Next 'i
					Call saveToCache("ssCategorySearch" & pstrCacheName, paryCategories, DateAdd("s", plngCacheTime, Now()))
					.MoveFirst
				End If
			End With
			Call DebugRecordTime("Categories loaded")
			'Call DebugPrintRecordset_Complete("Category Records", prsCategory)
		Else
			If cblnDebugCategorySearchTool Then Response.Write "<b>Loading Categories from application cache</b><br />"
			paryCategories = getFromCache("ssCategorySearch")
		End If
		
		If pblnCategoriesLoaded Then 
			prsCategory.MoveFirst
			Exit Sub
		End If
		
		pblnCategoriesLoaded = True
			
		If cblnDebugCategorySearchTool Then
			Response.Write "<fieldset><legend>LoadCategories</legend>"
			Response.Write "pblnCategoriesLoaded = " & pblnCategoriesLoaded & "<br />"
			Response.Write "</fieldset>"
			Response.Flush
		End If

	End Sub	'LoadCategories


	'***********************************************************************************************

	Public Function LoadCategories_Basic()
	
	Dim i
	Dim pstrSQL
	
		If cblnDebugCategorySearchTool Then
			Response.Write "<fieldset><legend>LoadCategories_Basic</legend>"
			Response.Write "pblnCategoriesLoaded = " & pblnCategoriesLoaded & "<br />"
			Response.Write "</fieldset>"
			Response.Flush
		End If
		
		If pblnCategoriesLoaded Then 
			prsCategory.MoveFirst
			Exit Function
		End If

		pstrSQL = "SELECT catID, catName, catDescription, catImage, catHttpAdd" _
				& " FROM sfCategories Where catIsActive=1" _
				& " ORDER BY catName"

		'Response.Write "pstrSQL = " & pstrSQL & "<br />"
		Set prsCategory = CreateObject("ADODB.RECORDSET")
		With prsCategory
			.ActiveConnection = cnn
			.CursorLocation = 2 'adUseClient
			.CursorType = 3 'adOpenStatic
			.LockType = 1 'adLockReadOnly
			.Source = pstrSQL
			.Open

			If cblnDebugCategorySearchTool Then
				Response.Write "<fieldset><legend>Load LoadCategories_Basic</legend>"
				Response.Write "State = " & .State & "<br />"
				Response.Write "EOF = " & .EOF & "<br />"
				Response.Write "RecordCount = " & .RecordCount & "<br />"
				Response.Write "</fieldset>"
			End If
			Call removeFromCache("ssCategorySearch")
			If Not isCacheItemExpired("ssCategorySearch") Then
				paryCategories = getFromCache("ssCategorySearch")
			End If
			If Not isArray(paryCategories) Then
				If Not .EOF Then
					ReDim paryCategories(.RecordCount - 1, en_CatFields_NumFields)
					For i = 0 To .RecordCount - 1
						'Response.Write i & ": " &  .Fields("catName").Value & "<br />"
						paryCategories(encatID, i) = .Fields("catID").Value
						paryCategories(encatName, i) = Trim(.Fields("catName").Value & "")
						paryCategories(encatHttpAdd, i) = Trim(.Fields("catHttpAdd").Value & "")

						'paryCategories(i, en_CatFields_uid) = .Fields("catID").Value
						'paryCategories(i, en_CatFields_ParentLevel) = .Fields("catID").Value
						'paryCategories(i, en_CatFields_Name) = .Fields("catName").Value
						'paryCategories(i, en_CatFields_ParentID) = .Fields("catID").Value
						paryCategories(i, en_CatFields_IsActive) = 1
						'paryCategories(i, en_CatFields_Description) = .Fields("catDescription").Value
						'paryCategories(i, en_CatFields_URL) = .Fields("catHttpAdd").Value
						'paryCategories(i, en_CatFields_ImagePath) = .Fields("catImage").Value
						'paryCategories(i, en_CatFields_CategoryID) = .Fields("catID").Value
						paryCategories(i, en_CatFields_IsBottom) = 1
						paryCategories(i, en_CatFields_InTrail) = 0
						.MoveNext
					Next 'i
					.MoveFirst
					
					'Save to Application
					If isArray(paryCategories) Then
						Call saveToCache("ssCategorySearch", paryCategories, DateAdd("s", 600, Now()))
					End If
				End If

			End If
		End With
		
		pblnCategoriesLoaded = True
			
	End Function	'LoadCategories_Basic

	'***********************************************************************************************

	Public Sub LoadCategoryScript

	Dim pstrCatName, pstrSubCatName, pstrSubSubCatName
	Dim plngCatCount, plngSubCatCount, plngSubSubCatCount
	Dim pblnSubOpen, pblnSubSubOpen
	Dim i

		Call LoadCategories
		With prsCategory
			pblnSubOpen = False
			pblnSubSubOpen = False
			plngCatCount = -1
			plngSubCatCount = -1
			plngSubSubCatCount = -1
			For i = 1 to .RecordCount
				If .Fields("catName").Value <> pstrCatName Then
					pstrCatOptions = pstrCatOptions & "<option value='" & Trim(.Fields("catID").Value) & "'>" & Server.HTMLEncode(Trim(.Fields("catName").Value)) & "</option>" & vbcrlf
					pstrCatName = .Fields("catName").Value
					plngCatCount = plngCatCount + 1
					If pblnSubOpen Then
						pstrSubCatOptions = pstrSubCatOptions & chr(34) & ";" & vbcrlf
						pblnSubOpen = False
						pstrSubCatName = ""
					End If
					If pblnSubSubOpen Then
						pstrSubSubCatOptions = pstrSubSubCatOptions & chr(34) & ";" & vbcrlf
						pblnSubSubOpen = False
						pstrSubSubCatName = ""
					End If

				End If

				If (Trim(paryCategories(enmfgName, i)) <> Trim(pstrSubCatName)) Then
					If pblnSubSubOpen Then
						pstrSubSubCatOptions = pstrSubSubCatOptions & chr(34) & ";" & vbcrlf
						pblnSubSubOpen = False
						pstrSubSubCatName = ""
					End If

					If Not pblnSubOpen Then
						pstrSubCatOptions = pstrSubCatOptions & "marySub[" & Trim(.Fields("catID").Value) & "] = " & chr(34) & "ALL|" & pstrSubCatText
						pblnSubOpen = True
					End If
					pstrSubCatName = paryCategories(enmfgName, i)
					plngSubCatCount = plngSubCatCount + 1
					If (paryCategories(enmfgID, i) <> 1) Then pstrSubCatOptions = pstrSubCatOptions & "|" & Trim(paryCategories(enmfgID, i)) & "|"  & Trim(paryCategories(enmfgName, i))
					
					If Len(pstrSubSubCatDBField) > 0 Then
						If paryCategories(envendName, i) <> pstrSubSubCatName Then
							If Not pblnSubSubOpen Then
								pstrSubSubCatOptions = pstrSubSubCatOptions & "marySubSub[" & chr(34) & Trim(.Fields("catID").Value) & "," & Trim(paryCategories(enmfgID, i)) & chr(34) & "] = " & chr(34) & "ALL|" & pstrSubSubCatText
								pblnSubSubOpen = True
							End If
							pstrSubSubCatName = paryCategories(envendName, i)
							plngSubSubCatCount = plngSubSubCatCount + 1
							If paryCategories(envendID, i) <> 1 Then pstrSubSubCatOptions = pstrSubSubCatOptions & "|" & Trim(paryCategories(envendID, i)) & "|"  & Trim(paryCategories(envendName, i))
						End If
					End If
				Else
					If Len(pstrSubSubCatDBField) > 0 Then
						If paryCategories(envendName, i) <> pstrSubSubCatName AND paryCategories(envendID, i) <> 1 Then
							pstrSubSubCatName = paryCategories(envendName, i)
							plngSubSubCatCount = plngSubSubCatCount + 1
							If paryCategories(envendID, i) <> 1 Then pstrSubSubCatOptions = pstrSubSubCatOptions & "|" & Trim(paryCategories(envendID, i)) & "|"  & Trim(paryCategories(envendName, i))
						End If
					End If
				End If
				.MoveNext
			Next
			
			If pblnSubSubOpen Then
				pstrSubSubCatOptions = pstrSubSubCatOptions & chr(34) & ";" & vbcrlf
				pblnSubSubOpen = False
				pstrSubSubCatName = ""
			End If
			If pblnSubOpen Then
				pstrSubCatOptions = pstrSubCatOptions & chr(34) & ";" & vbcrlf
				pblnSubOpen = False
				pstrSubCatName = ""
			End If
			
		End With
		
	End Sub 'LoadCategoryScript

	'***********************************************************************************************

	Public Sub WriteSearchForm

		With Response
			.Write "<form id=frmCategory name=frmCategory>" & vbcrlf
			.Write "<table border=1 cellspacing=0 cellpadding=3>" & vbcrlf
			.Write "<tr><td>" & vbcrlf
			.Write "<table border=0 cellspacing=0 cellpadding=3 width='100%'>" & vbcrlf
			.Write "<tr>" & vbcrlf
			.Write "<td colspan=4>" & vbcrlf
			.Write "<table border='0' cellspacing=0 cellpadding=3 width='100%'>" & vbcrlf
			.Write "<tr>" & vbcrlf
			.Write "<td>" & vbcrlf
			.Write "<b>Search</b>&nbsp;<input name='txtsearchParamTxt' value=''/>" & vbcrlf
			.Write "<br /><input type='radio' value='ALL' checked name='txtsearchParamType'>&nbsp;<b>ALL</b>&nbsp;Words&nbsp;<input type='radio' name='txtsearchParamType' value='ANY'>&nbsp;<b>ANY</b>&nbsp;Words&nbsp;<input type='radio' name='txtsearchParamType' value='Exact'>&nbsp;Exact&nbsp;Phrase" & vbcrlf
			.Write "</td>" & vbcrlf
			.Write "<td><a href='' onclick='Search(); return false;'>Search</a></td>" & vbcrlf
			.Write "</tr>" & vbcrlf
			.Write "</table>" & vbcrlf
			.Write "</td>" & vbcrlf
			.Write "</tr>" & vbcrlf
			.Write "<tr>" & vbcrlf
			.Write "<tr><td colspan=3>narrow your search by category (optional)</td></tr>" & vbcrlf
			.Write "<tr>" & vbcrlf
			.Write "<td>" & vbcrlf
			WriteCategory "selCat",10
			.Write "</td>" & vbcrlf
			.Write "<td>" & vbcrlf
			WriteSubCategory "selMfg",10
			.Write "</td>" & vbcrlf
			.Write "<td>" & vbcrlf
			WriteSubSubCategory "selVend",10
			.Write "</td>" & vbcrlf
			.Write "</tr>" & vbcrlf
			.Write "</table>" & vbcrlf
			.Write "</td></tr></table>" & vbcrlf
			.Write "</form>" & vbcrlf
		End With
		WriteChangeScript "frmCategory" 

	End Sub	'WriteSearchForm

	'***********************************************************************************************

	Public Sub WriteCategoryScript

		pstrFormName = "frmCategory"
		Response.Write "<form id='frmCategory' name='frmCategory' action='' method='GET'>" & vbcrlf
		Call WriteCategory("selCat",10)
		Call WriteSubCategory("selMfg",10)
		Call WriteSubSubCategory("selVend",10)
		Call WriteChangeScript
		Response.Write "</form>" & vbcrlf
		
	End Sub 'WriteCategoryScript

	'***********************************************************************************************

	Public Sub WriteChangeScript(strFormName)

		pstrFormName = strFormName
		With Response
			.Write "<script language='javascript1.3'>" & vbcrlf
			.Write "var marySub = new Array();" & vbcrlf
			.Write "var marySubSub = new Array();" & vbcrlf
			.Write "" & vbcrlf
			.Write pstrSubCatOptions
			.Write "" & vbcrlf
			.Write pstrSubSubCatOptions
			.Write "" & vbcrlf
		
			.Write "" & vbcrlf
			.Write "var mlngCatID=1;" & vbcrlf
			.Write "var mlngMfgID=1;" & vbcrlf
			.Write "var mlngVendID=1;" & vbcrlf
			.Write "" & vbcrlf
			.Write "var mblnIsNS = !(document.all);" & vbcrlf
			.Write "function clearOptions(theSelect){for (var i=theSelect.options.length-1; i >= 0;i--){if (mblnIsNS){theSelect.options[i].remove;}else{theSelect.remove(i);}}theSelect.length=0;}" & vbcrlf
			.Write "function getSelectValue(theSelect){if (theSelect.selectedIndex == -1){return('');}else{return(theSelect.options[theSelect.selectedIndex].value);}}" & vbcrlf
			.Write "" & vbcrlf
			.Write "function selectCat(){" & vbcrlf
			.Write "    mlngCatID = getSelectValue(mtheForm." & pstrCatField & ");" & vbcrlf
			.Write "	var parySub = (marySub[mlngCatID]).split(" & chr(34) & "|" & chr(34) & ");" & vbcrlf
			.Write "	clearOptions(mtheForm." & pstrSubCatField & ");" & vbcrlf
			.Write "	for (var i=0; i < parySub.length/2;i++){mtheForm." & pstrSubCatField & ".options[i] = new Option();mtheForm." & pstrSubCatField & ".options[i].value = parySub[i*2];mtheForm." & pstrSubCatField & ".options[i].text = parySub[i*2+1];}" & vbcrlf
			.Write "	mtheForm." & pstrSubCatField & "[0].selected = true;" & vbcrlf
			.Write "	selectSub(mtheForm);" & vbcrlf
			.Write "}" & vbcrlf
			.Write "" & vbcrlf

			.Write "function selectSub(){" & vbcrlf
			.Write "	mlngMfgID = getSelectValue(mtheForm." & pstrSubCatField & ");" & vbcrlf
			.Write "" & vbcrlf
			.Write "	var parySubSub;" & vbcrlf
			.Write "	var pstrTemp = mlngCatID + ',' + mlngMfgID;" & vbcrlf
			.Write "	var pstrTemp2 = marySubSub[pstrTemp];" & vbcrlf
			
			.Write "	if((typeof(pstrTemp2) != 'undefined'))" & vbcrlf
			.Write "	{" & vbcrlf
			.Write "	    parySubSub = (marySubSub[pstrTemp]).split(" & chr(34) & "|" & chr(34) & ");" & vbcrlf
			.Write "	}else{" & vbcrlf
			.Write "	    pstrTemp2 = " & chr(34) & "ALL|" & pstrSubSubCatText & chr(34) & ";" & vbcrlf
			.Write "	    parySubSub = pstrTemp2.split(" & chr(34) & "|" & chr(34) & ");" & vbcrlf
			.Write "	}" & vbcrlf
			.Write "" & vbcrlf
			.Write "	clearOptions(mtheForm." & pstrSubSubCatField & ");" & vbcrlf
			.Write "	for (var i=0; i < parySubSub.length/2;i++)" & vbcrlf
			.Write "	{mtheForm." & pstrSubSubCatField & ".options[i] = new Option();mtheForm." & pstrSubSubCatField & ".options[i].value = parySubSub[i*2];mtheForm." & pstrSubSubCatField & ".options[i].text = parySubSub[i*2+1];}" & vbcrlf
			.Write "	mtheForm." & pstrSubSubCatField & "[0].selected = true;" & vbcrlf
			.Write "}" & vbcrlf
			.Write "" & vbcrlf
			.Write "function selectSubSub(){" & vbcrlf
			.Write "	mlngVendID = getSelectValue(mtheForm." & pstrSubSubCatField & ");" & vbcrlf
			.Write "}" & vbcrlf
			.Write "" & vbcrlf

			.Write "function Search(){" & vbcrlf
			.Write "	mlngCatID = getSelectValue(mtheForm." & pstrCatField & ");" & vbcrlf
			.Write "	mlngMfgID = getSelectValue(mtheForm." & pstrSubCatField & ");" & vbcrlf
			.Write "	mlngVendID = getSelectValue(mtheForm." & pstrSubSubCatField & ");" & vbcrlf
			.Write "	for (var i=0; i < mtheForm.txtsearchParamType.length;i++){" & vbcrlf
			.Write "		if (mtheForm.txtsearchParamType[i].checked){" & vbcrlf
			.Write "			var ptxtsearchParamType = mtheForm.txtsearchParamType[i].value;" & vbcrlf
			.Write "			break;" & vbcrlf
			.Write "		}" & vbcrlf
			.Write "	}" & vbcrlf
			.Write "	window.location = 'search_results.asp?&txtsearchParamType=' + ptxtsearchParamType + '&txtsearchParamTxt=' + mtheForm.txtsearchParamTxt.value + '&txtsearchParamCat=' + mlngCatID + '&txtsearchParamMan=' + mlngMfgID + '&txtsearchParamVen=' + mlngVendID;" & vbcrlf
			.Write "}" & vbcrlf
			.Write "" & vbcrlf
		
			.Write "var mtheForm = document." & pstrFormName & ";" & vbcrlf
			.Write "mtheForm." & pstrCatField & "[0].selected = true;" & vbcrlf
			.Write "    mlngCatID = mtheForm." & pstrCatField & "[0].value;" & vbcrlf
			.Write "selectCat();" & vbcrlf
			.Write "</script>" & vbcrlf
		End With
		
	End Sub	'WriteChangeScript

	'***********************************************************************************************

	Public Sub WriteCategory(strFieldName,intSize)
		pstrCatField = strFieldName
		Response.Write "<select id=" & strFieldName & " name=" & strFieldName & " onchange='return selectCat();' size='" & intSize & "'>"  & vbcrlf & pstrCatOptions & "</select>" & vbcrlf
	End Sub
	Public Sub WriteSubCategory(strFieldName,intSize)
		pstrSubCatField = strFieldName
		Response.Write "<select id=" & strFieldName & " name=" & strFieldName & " onchange='return selectSub();' size='" & intSize & "'><option value='ALL'>" & Server.HTMLEncode(pstrSubCatText) & "</option></select>" & vbcrlf
	End Sub
	Public Sub WriteSubSubCategory(strFieldName,intSize)
		pstrSubSubCatField = strFieldName
		Response.Write "<select id=" & strFieldName & " name=" & strFieldName & " onchange='return selectSubSub();' size='" & intSize & "'><option value='ALL'>" & Server.HTMLEncode(pstrSubSubCatText) & "</option></select>" & vbcrlf
	End Sub

	'***********************************************************************************************

	Public Sub WriteManufacturersAsTable

	Dim i
	Dim pclsSFSearch
	Dim plngNumItems
	Dim pstrLineOut
	Dim pstrLink
	Dim pstrTemp
	
		pstrTemp = pstrTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf

		If loadManufacturerArray Then
			plngNumItems = UBound(maryManufacturers)
			For i = 0 To plngNumItems
				pstrLineOut = ""
				If maryManufacturers(i)(0) <> 1 And Not maryManufacturers(i)(9) Then	'Exclude No Manufacturer
					Set pclsSFSearch = New sfSearch
					pclsSFSearch.txtsearchParamMan = maryManufacturers(i)(0)
					'Response.Write maryManufacturers(i)(1) & "<br />"
					pstrLink = "<a href=" & Chr(34) & "search_results.asp" & "?" & pclssfSearch.SearchLinkParameters(False) & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & Server.HTMLEncode(maryManufacturers(i)(1)) & pstrEndFont & "</a>"
					Set pclsSFSearch = Nothing
					pstrLineOut = "<tr><td class=""" & pstrTDClass & """ valign=""middle"">" & pstrLink & "</td></tr>" & vbcrlf
					pstrTemp = pstrTemp & pstrLineOut
				End If
			Next 'i
		End If
		pstrTemp = pstrTemp & "</table>" & vbcrlf
		
		Response.Write pstrTemp
			
	End Sub	'WriteManufacturersAsTable

	'***********************************************************************************************

	Private Sub OpenItem(ByVal strDBField, ByRef strName, ByRef blnOpen, ByVal bytLevel)
		If Not blnOpen Then
			Select Case bytLevel
				Case 1:	Response.Write cstrStartCat & vbcrlf
				Case 2:	Response.Write cstrStartSubCat & vbcrlf
				Case 3:	Response.Write cstrStartSubSubCat & vbcrlf
				Case 4:	Response.Write cstrStartProduct & vbcrlf
			End Select
			strName = strDBField
			blnOpen = True
		End If
	End Sub	'OpenItem

	Private Sub CloseItem(ByRef strName, ByRef blnOpen, ByVal bytLevel)
		If blnOpen Then
			Select Case bytLevel
				Case 1:	Response.Write cstrEndCat & vbcrlf
				Case 2:	Response.Write cstrEndSubCat & vbcrlf
				Case 3:	Response.Write cstrEndSubSubCat & vbcrlf
			End Select
			strName = ""
			blnOpen = False
		End If
	End Sub	'CloseItem
    
	Private Sub WriteItem(ByRef objRS, ByVal bytLevel)
	
	Dim pstrTempURL
	Dim pstrTempAdjustHTTPField	'necessary due to inconsistency in SF naming convention
	
		With objRS
		Select Case bytLevel
			Case 1:	
				If Len(.Fields("catHttpAdd").Value & "") > 0 Then
					pstrTempURL = .Fields("catHttpAdd").Value
				Else
					pstrTempURL = "search_results.asp?&txtsearchParamType=ALL&txtsearchParamCat=" & Trim(.Fields("catID").Value) & "&txtsearchParamMan=ALL&txtsearchParamVen=ALL'"
				End If
				If (pstrCurrentCategory = CStr(Trim(.Fields("catID").Value & ""))) And pblnHighlightSelected Then
					Response.Write cstrStartCatDetail & "<a " & cstrCatLinkStyle_Selected & " href='" & pstrTempURL & "'>" & Trim(.Fields("catName").Value) & "</a>" & cstrEndCatDetail & vbcrlf
				Else
					Response.Write cstrStartCatDetail & "<a " & cstrCatLinkStyle & " href='" & pstrTempURL & "'>" & Trim(.Fields("catName").Value) & "</a>" & cstrEndCatDetail & vbcrlf
				End If
			Case 2:	
				If pstrSubCatDBField = "vend" Then
					pstrTempAdjustHTTPField = "HttpAddr"
				Else
					pstrTempAdjustHTTPField = "HttpAdd"
				End If
				If Len(.Fields(pstrSubCatDBField & pstrTempAdjustHTTPField).Value & "") > 0 Then
					pstrTempURL = .Fields(pstrSubCatDBField & pstrTempAdjustHTTPField).Value
				Else
					pstrTempURL = "search_results.asp?&txtsearchParamType=ALL&txtsearchParamCat=" & Trim(.Fields("catID").Value) & "&txtsearchParamMan=" & Trim(.Fields(pstrSubCatDBField & "ID").Value) & "&txtsearchParamVen=ALL"
				End If
				If .Fields(pstrSubCatDBField & "ID").Value <> 1 Then 
					If (pstrCurrentSubCategory = CStr(Trim(.Fields(pstrSubCatDBField & "ID").Value & ""))) And pblnHighlightSelected Then
						Response.Write "    <li " & cstrSubCatLIStyle & "><a " & cstrSubCatLinkStyle_Selected & " href='" & pstrTempURL & "'>" & Trim(.Fields(pstrSubCatDBField & "Name").Value) & "</a></li>" & vbcrlf
					Else
						Response.Write "    <li " & cstrSubCatLIStyle & "><a " & cstrSubCatLinkStyle & " href='" & pstrTempURL & "'>" & Trim(.Fields(pstrSubCatDBField & "Name").Value) & "</a></li>" & vbcrlf
					End If
				End If
			Case 3:	
				If pstrSubSubCatDBField = "vend" Then
					pstrTempAdjustHTTPField = "HttpAddr"
				Else
					pstrTempAdjustHTTPField = "HttpAdd"
				End If
				If Len(.Fields(pstrSubSubCatDBField & pstrTempAdjustHTTPField).Value & "") > 0 Then
					pstrTempURL = .Fields(pstrSubSubCatDBField & pstrTempAdjustHTTPField).Value
				Else
					pstrTempURL = "search_results.asp?&txtsearchParamType=ALL&txtsearchParamCat=" & Trim(.Fields("catID").Value) & "&txtsearchParamMan=" & Trim(.Fields(pstrSubCatDBField & "ID").Value) & "&txtsearchParamVen=" & Trim(.Fields(pstrSubSubCatDBField & "ID").Value)
				End If
				If .Fields(pstrSubSubCatDBField & "ID").Value <> 1 Then 
					If (pstrCurrentSubSubCategory = CStr(Trim(.Fields(pstrSubSubCatDBField & "ID").Value & ""))) And pblnHighlightSelected Then
						Response.Write "      <li " & cstrSubSubCatLIStyle & "><a " & cstrSubSubCatLinkStyle_Selected & " href='" & pstrTempURL & "'>" & Trim(.Fields(pstrSubSubCatDBField & "Name").Value) & "</a></li>" & vbcrlf
					Else
						Response.Write "      <li " & cstrSubSubCatLIStyle & "><a " & cstrSubSubCatLinkStyle & " href='" & pstrTempURL & "'>" & Trim(.Fields(pstrSubSubCatDBField & "Name").Value) & "</a></li>" & vbcrlf
					End If
				End If
			Case 4:	
				Response.Write cstrStartProduct & vbcrlf
		End Select
		End With
	End Sub	'WriteItem

	'***********************************************************************************************

	Private Sub WriteArrayItem(ByVal lngRecord, ByVal bytLevel)
	
	Dim pstrTempURL
	Dim pstrTempAdjustHTTPField	'necessary due to inconsistency in SF naming convention
	
		Select Case bytLevel
			Case 1:	
				If Len(paryCategories(encatHttpAdd, lngRecord) & "") > 0 Then
					pstrTempURL = paryCategories(encatHttpAdd, lngRecord)
				Else
					pstrTempURL = "search_results.asp?&txtsearchParamType=ALL&txtsearchParamCat=" & Trim(paryCategories(encatID, lngRecord)) & "&txtsearchParamMan=ALL&txtsearchParamVen=ALL'"
				End If
				If (pstrCurrentCategory = CStr(Trim(paryCategories(encatID, lngRecord) & ""))) And pblnHighlightSelected Then
					Response.Write cstrStartCatDetail & "<a " & cstrCatLinkStyle_Selected & " href='" & pstrTempURL & "'>" & Trim(paryCategories(encatName, lngRecord)) & "</a>" & cstrEndCatDetail & vbcrlf
				Else
					Response.Write cstrStartCatDetail & "<a " & cstrCatLinkStyle & " href='" & pstrTempURL & "'>" & Trim(paryCategories(encatName, lngRecord)) & "</a>" & cstrEndCatDetail & vbcrlf
				End If
			Case 2:	
				If pstrSubCatDBField = "vend" Then
					pstrTempAdjustHTTPField = "HttpAddr"
				Else
					pstrTempAdjustHTTPField = "HttpAdd"
				End If
				If Len(paryCategories(enmfgHttpAdd, lngRecord) & "") > 0 Then
					pstrTempURL = paryCategories(enmfgHttpAdd, lngRecord)
				Else
					pstrTempURL = "search_results.asp?&txtsearchParamType=ALL&txtsearchParamCat=" & Trim(paryCategories(encatID, lngRecord)) & "&txtsearchParamMan=" & Trim(paryCategories(enmfgID, lngRecord)) & "&txtsearchParamVen=ALL"
				End If
				If paryCategories(enmfgID, lngRecord) <> 1 Then 
					If (pstrCurrentSubCategory = CStr(Trim(paryCategories(enmfgID, lngRecord) & ""))) And pblnHighlightSelected Then
						Response.Write "    <li " & cstrSubCatLIStyle & "><a " & cstrSubCatLinkStyle_Selected & " href='" & pstrTempURL & "'>" & Trim(paryCategories(enmfgName, lngRecord)) & "</a></li>" & vbcrlf
					Else
						Response.Write "    <li " & cstrSubCatLIStyle & "><a " & cstrSubCatLinkStyle & " href='" & pstrTempURL & "'>" & Trim(paryCategories(enmfgName, lngRecord)) & "</a></li>" & vbcrlf
					End If
				End If
			Case 3:	
				If pstrSubSubCatDBField = "vend" Then
					pstrTempAdjustHTTPField = "HttpAddr"
				Else
					pstrTempAdjustHTTPField = "HttpAdd"
				End If
				If Len(paryCategories(envendHttpAddr, lngRecord) & "") > 0 Then
					pstrTempURL = paryCategories(envendHttpAddr, lngRecord)
				Else
					pstrTempURL = "search_results.asp?&txtsearchParamType=ALL&txtsearchParamCat=" & Trim(paryCategories(encatID, lngRecord)) & "&txtsearchParamMan=" & Trim(paryCategories(enmfgID, lngRecord)) & "&txtsearchParamVen=" & Trim(paryCategories(envendID, lngRecord))
				End If
				If paryCategories(envendID, lngRecord) <> 1 Then 
					If (pstrCurrentSubSubCategory = CStr(Trim(paryCategories(envendID, lngRecord) & ""))) And pblnHighlightSelected Then
						Response.Write "      <li " & cstrSubSubCatLIStyle & "><a " & cstrSubSubCatLinkStyle_Selected & " href='" & pstrTempURL & "'>" & Trim(paryCategories(envendName, lngRecord)) & "</a></li>" & vbcrlf
					Else
						Response.Write "      <li " & cstrSubSubCatLIStyle & "><a " & cstrSubSubCatLinkStyle & " href='" & pstrTempURL & "'>" & Trim(paryCategories(envendName, lngRecord)) & "</a></li>" & vbcrlf
					End If
				End If
			Case 4:	
				Response.Write cstrStartProduct & vbcrlf
		End Select

	End Sub	'WriteArrayItem

	'***********************************************************************************************

	Public Sub WriteCategories()

	Dim i
	Dim pblnSubOpen
	Dim pblnSubSubOpen
	Dim plngNumCategories
	Dim pstrCatName
	Dim pstrSubCatName
	Dim pstrSubSubCatName

		'On Error Resume Next
		On Error Goto 0
		
		If cblnDebugCategorySearchTool Then
			Response.Write "<fieldset><legend>WriteCategories</legend>"
			Response.Write "pblnDisplayTopCategoriesOnly: " & pblnDisplayTopCategoriesOnly & "<br />"
			Response.Write "pblnDisplayCurrentCategoryOnly: " & pblnDisplayCurrentCategoryOnly & "<hr>"
			Response.Write "pstrCurrentCategory: " & pstrCurrentCategory & "<br />"
			Response.Write "pstrCurrentSubCategory: " & pstrCurrentSubCategory & "<br />"
			Response.Write "pstrCurrentSubSubCategory: " & pstrCurrentSubSubCategory & "<br />"
			Response.Write "</fieldset>"
			Response.Flush
		End If
		
		If pblnDisplayTopCategoriesOnly Then
			Call LoadCategories_Basic
			pstrSubCatDBField = ""
			pstrSubSubCatDBField = ""
		Else
			Call LoadCategories
		End If

		plngNumCategories = UBound(paryCategories, 1)
		pblnSubOpen = False
		pblnSubSubOpen = False
		Response.Write cstrStartCat & vbcrlf
		For i = 0 to plngNumCategories
			If paryCategories(encatName, i) <> pstrCatName Then
				Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
				Call CloseItem(pstrSubCatName, pblnSubOpen, 2)
				Call WriteArrayItem(i, 1)
				If (pstrCurrentCategory = CStr(Trim(paryCategories(encatID, i) & ""))) Or ((Len(pstrCurrentCategory) = 0) And Not pblnDisplayCurrentCategoryOnly) Then
					If Len(pstrSubCatDBField) > 0 Then
						If paryCategories(enmfgName, i) <> pstrSubCatName Then
							Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
							If Not pblnSubOpen Then
								Response.Write cstrStartSubCat & vbcrlf
								pblnSubOpen = True
							End If
							Call WriteArrayItem(i,2)
							If Len(pstrSubSubCatDBField) > 0 Then
								If paryCategories(envendName, i) <> pstrSubSubCatName Then
									Call OpenItem(paryCategories(envendName, i), pstrSubSubCatName, pblnSubSubOpen, 3)
									Call WriteArrayItem(i, 3)
								End If
							End If
							pstrSubCatName = paryCategories(enmfgName, i)
						Else
							If Len(pstrSubSubCatDBField) > 0 Then
								If paryCategories(envendName, i) <> pstrSubSubCatName Then
									Call OpenItem(paryCategories(envendName, i), pstrSubSubCatName, pblnSubSubOpen, 3)
									Call WriteArrayItem(i, 3)
								End If
							End If
						End If	'paryCategories(enmfgName, i) <> pstrSubCatName
					End If	'Len(pstrSubCatDBField) > 0
				End If
				pstrCatName = paryCategories(encatName, i)
			Else
				If (pstrCurrentCategory = CStr(Trim(paryCategories(encatID, i) & ""))) Or ((Len(pstrCurrentCategory) = 0) And Not pblnDisplayCurrentCategoryOnly) Then
					If Len(pstrSubCatDBField) > 0 Then
						If paryCategories(enmfgName, i) <> pstrSubCatName Then
							Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
							If Not pblnSubOpen Then
								Response.Write cstrStartSubCat & vbcrlf
								pblnSubOpen = True
							End If
							Call WriteArrayItem(i,2)
							If Len(pstrSubSubCatDBField) > 0 Then
								If paryCategories(envendName, i) <> pstrSubSubCatName Then
									Call OpenItem(paryCategories(envendName, i), pstrSubSubCatName, pblnSubSubOpen, 3)
									Call WriteArrayItem(i, 3)
									pblnSubOpen = True
								End If
							End If
							pstrSubCatName = paryCategories(enmfgName, i)
						Else
							If Len(pstrSubSubCatDBField) > 0 Then
								If paryCategories(envendName, i) <> pstrSubSubCatName Then
									Call OpenItem(paryCategories(envendName, i), pstrSubSubCatName, pblnSubSubOpen, 3)
									Call WriteArrayItem(i, 3)
								End If
							End If
						End If	'paryCategories(enmfgName, i) <> pstrSubCatName
					End If	'Len(pstrSubCatDBField) > 0
				End If
			End If
		Next
		Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
		Call CloseItem(pstrSubCatName, pblnSubOpen, 2)
		Response.Write cstrEndCat & vbcrlf

	End Sub	'WriteCategories

	'***********************************************************************************************

	Public Sub WriteProducts

	Dim pstrSQL
	Dim pstrSQLOrderBy
	Dim prsProducts
	Dim pstrCatName, pstrSubCatName, pstrSubSubCatName
	Dim pblnSubOpen, pblnSubSubOpen, pblnProdOpen
	Dim pstrProdLink
	Dim i

		pstrSQLOrderBy = " ORDER BY sfCategories.catName"
		If Len(pstrSubCatDBField) > 0 Then pstrSQLOrderBy = pstrSQLOrderBy & ", " & pstrSubCatDBField& "Name"
		If Len(pstrSubSubCatDBField) > 0 Then pstrSQLOrderBy = pstrSQLOrderBy & ", " & pstrSubSubCatDBField& "Name"

		'Response.Write "pstrSQL = " & pstrSQL & "<br />"
		pstrSQL = "SELECT sfCategories.catID, sfCategories.catName, sfManufacturers.mfgID, sfManufacturers.mfgName, sfVendors.vendID, sfVendors.vendName, sfProducts.prodID, sfProducts.prodName, sfProducts.prodShortDescription, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfCategories.catHttpAdd, sfManufacturers.mfgHttpAdd, sfVendors.vendHttpAddr " _
				& " FROM sfCategories INNER JOIN (sfVendors INNER JOIN (sfManufacturers INNER JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) ON sfVendors.vendID = sfProducts.prodVendorId) ON sfCategories.catID = sfProducts.prodCategoryId" _
				& " WHERE sfProducts.prodEnabledIsActive=1" _
				& pstrSQLOrderBy & ", sfProducts.prodID"

		'Response.Write "<h3>" & pstrSQL & "</h3>"
		Set prsProducts = CreateObject("ADODB.RECORDSET")
		With prsProducts
			.ActiveConnection = cnn
		    .CursorLocation = 2 'adUseClient
		    .CursorType = 3 'adOpenStatic
		    .LockType = 1 'adLockReadOnly
			.Source = pstrSQL
			.Open
		End With

		With prsProducts
			
			pblnSubOpen = False
			pblnSubSubOpen = False
			pblnProdOpen = False
			
			Response.Write cstrStartCat & vbcrlf
			
			For i = 1 to .RecordCount
				If .Fields("catName").Value <> pstrCatName Then
				
					If pblnProdOpen Then
						Response.Write cstrEndProduct & vbcrlf
						pblnProdOpen = False
					End If
					Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
					Call CloseItem(pstrSubCatName, pblnSubOpen, 2)
					
					Call WriteItem(prsProducts, 1)
					If .Fields(pstrSubCatDBField & "Name").Value <> pstrSubCatName Then
						If pblnProdOpen Then
							Response.Write cstrEndProduct & vbcrlf
							pblnProdOpen = False
						End If
						Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
						If Not pblnSubOpen Then
							Response.Write cstrStartSubCat & vbcrlf
							pblnSubOpen = True
						End If
						Call WriteItem(prsProducts, 2)
						If Len(pstrSubSubCatDBField) > 0 Then
							If .Fields(pstrSubSubCatDBField & "Name").Value <> pstrSubSubCatName Then
								Call OpenItem(.Fields(pstrSubSubCatDBField & "Name").Value, pstrSubSubCatName, pblnSubSubOpen, 3)
								Call WriteItem(prsProducts, 3)
							End If
						End If
						pstrSubCatName = .Fields(pstrSubCatDBField & "Name").Value
					Else
						If Len(pstrSubSubCatDBField) > 0 Then
							If .Fields(pstrSubSubCatDBField & "Name").Value <> pstrSubSubCatName Then
								If pblnProdOpen Then
									Response.Write cstrEndProduct & vbcrlf
									pblnProdOpen = False
								End If
								Call OpenItem(.Fields(pstrSubSubCatDBField & "Name").Value, pstrSubSubCatName, pblnSubSubOpen, 3)
								Call WriteItem(prsProducts, 3)
							End If
						End If
					End If
					pstrCatName = .Fields("catName").Value
				Else
					If .Fields(pstrSubCatDBField & "Name").Value <> pstrSubCatName Then
						If pblnProdOpen Then
							Response.Write cstrEndProduct & vbcrlf
							pblnProdOpen = False
						End If
						Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
						If Not pblnSubOpen Then
							Response.Write cstrStartSubCat & vbcrlf
							pblnSubOpen = True
						End If
						Call WriteItem(prsProducts, 2)
						If Len(pstrSubSubCatDBField) > 0 Then
							If .Fields(pstrSubSubCatDBField & "Name").Value <> pstrSubSubCatName Then
								If Not pblnSubSubOpen Then
									Response.Write cstrStartSubSubCat & vbcrlf
									pblnSubSubOpen = True
								End If
								Call WriteItem(prsProducts, 3)
								pstrSubSubCatName = .Fields(pstrSubSubCatDBField & "Name").Value
								pblnSubOpen = True
							End If
						End If
						pstrSubCatName = .Fields(pstrSubCatDBField & "Name").Value
					Else
						If Len(pstrSubSubCatDBField) > 0 Then
							If .Fields(pstrSubSubCatDBField & "Name").Value <> pstrSubSubCatName Then
								If pblnProdOpen Then
									Response.Write cstrEndProduct & vbcrlf
									pblnProdOpen = False
								End If
								Call OpenItem(.Fields(pstrSubSubCatDBField & "Name").Value, pstrSubSubCatName, pblnSubSubOpen, 3)
								Call WriteItem(prsProducts, 3)
							End If
						End If
					End If
				End If

				If Not pblnProdOpen Then
					Response.Write cstrStartProduct & vbcrlf
					pblnProdOpen = True
				End If
				
				'Determine the product link
				'available fields - more can be added via the SQL statement
				'
				'prodID, prodName, prodShortDescription, prodImageSmallPath, prodLink

				If Len(.Fields("prodLink").Value & "") = 0 Then
					pstrProdLink = "detail.asp?product_id=" & Trim(.Fields("prodID").Value)
				Else
					pstrProdLink = Trim(.Fields("prodLink").Value)
				End If
				
				If True Then	'Replace with False to show small product images
					pstrProdLink = "<a " & cstrProductLinkStyle & " href=" & Chr(34) & pstrProdLink & Chr(34) & ">" & .Fields("prodName").Value & "</a>"
				Else
					If Len(.Fields("prodLink").Value & "") > 0 Then
						pstrProdLink = "<a " & cstrProductLinkStyle & " href=" & Chr(34) & pstrProdLink & Chr(34) & "><img border=0 src=" & Chr(34) & .Fields("prodImageSmallPath").Value & Chr(34) & ">&nbsp;"& .Fields("prodName").Value & "</a>"
					Else
						pstrProdLink = "<a " & cstrProductLinkStyle & " href=" & Chr(34) & pstrProdLink & Chr(34) & ">" & .Fields("prodName").Value & "</a>"
					End If
				End If
				Response.Write "        <li " & cstrProductLIStyle & ">" & pstrProdLink & "</li>" & vbcrlf
				.MoveNext
			Next
			If pblnProdOpen Then Response.Write cstrEndProduct & vbcrlf	'for the product
			Call CloseItem(pstrSubSubCatName, pblnSubSubOpen, 3)
			Call CloseItem(pstrSubCatName, pblnSubOpen, 2)
			Response.Write cstrEndCat & vbcrlf
			
			.Close
		end with
		Set prsProducts = Nothing
		
	End Sub	'WriteProducts

	'***********************************************************************************************

	Sub WriteSingleSelect

	Dim pstrCatName, pstrSubCatName, pstrSubSubCatName
	Dim pstrSubCatIndent, pstrSubSubCatIndent
	Dim i
	Dim plngNumCategories

		pstrSubCatIndent = "&nbsp;&nbsp;-&nbsp;"
		pstrSubSubCatIndent = "&nbsp;&nbsp;&nbsp;&nbsp;--&nbsp;"
		
		Call LoadCategories
		plngNumCategories = UBound(paryCategories, 2)
		If plngNumCategories > 0 Then
			Response.Write "<input type='hidden' name='txtsearchParamCat' value='ALL'><select name='CatSource' id='CatSource' size='1' onchange='var theForm=this.form;var pstrTemp=this.value.split(""."");theForm.txtsearchParamCat.value=pstrTemp[0];theForm.txtsearchParamMan.value=pstrTemp[1];theForm.txtsearchParamVen.value=pstrTemp[2];' style='" & C_FORMDESIGN & "'><option value='ALL.ALL.ALL'>All " & C_CategoryNameP & "</option>"
			For i = 1 to plngNumCategories
				If paryCategories(encatName, i) <> pstrCatName Then
				Response.Write "  <option value='" & Trim(paryCategories(encatID, i)) & ".ALL.ALL'>" & Server.HTMLEncode(Trim(paryCategories(encatName, i))) & "</option>" & vbcrlf
					If paryCategories(enmfgName, i) <> pstrSubCatName Then
						If paryCategories(enmfgID, i) <> 1 Then Response.Write "    <option value='" & Trim(paryCategories(encatID, i)) & "." & Trim(paryCategories(enmfgID, i)) & ".ALL'>" & pstrSubCatIndent & Server.HTMLEncode(Trim(paryCategories(enmfgName, i))) & "</option>" & vbcrlf
						If Len(pstrSubSubCatDBField) > 0 Then
							If paryCategories(envendName, i) <> pstrSubSubCatName Then
								If paryCategories(envendID, i) <> 1 Then Response.Write "      <option value='" & Trim(paryCategories(encatID, i)) & "." & Trim(paryCategories(enmfgID, i)) & "." & Trim(paryCategories(envendID, i)) & "'>" & pstrSubSubCatIndent & Server.HTMLEncode(Trim(paryCategories(envendName, i))) & "</option>" & vbcrlf
								pstrSubSubCatName = paryCategories(envendName, i)
							End If
						End If
						pstrSubCatName = paryCategories(enmfgName, i)
					Else
						If Len(pstrSubSubCatDBField) > 0 Then
							If paryCategories(envendName, i) <> pstrSubSubCatName Then
								If paryCategories(envendID, i) <> 1 Then Response.Write "      <option value='" & Trim(paryCategories(encatID, i)) & "." & Trim(paryCategories(enmfgID, i)) & "." & Trim(paryCategories(envendID, i)) & "'>" & pstrSubSubCatIndent & Server.HTMLEncode(Trim(paryCategories(envendName, i))) & "</option>" & vbcrlf
								pstrSubSubCatName = paryCategories(envendName, i)
							End If
						End If
					End If
					pstrCatName = paryCategories(encatName, i)
				Else
					If paryCategories(enmfgName, i) <> pstrSubCatName Then
						If paryCategories(enmfgID, i) <> 1 Then Response.Write "    <option value='" & Trim(paryCategories(encatID, i)) & "." & Trim(paryCategories(enmfgID, i)) & ".ALL'>" & pstrSubCatIndent & Trim(paryCategories(enmfgName, i)) & "</option>" & vbcrlf
						If Len(pstrSubSubCatDBField) > 0 Then
							If paryCategories(envendName, i) <> pstrSubSubCatName Then
								If paryCategories(envendID, i) <> 1 Then Response.Write "      <option value='" & Trim(paryCategories(encatID, i)) & "." & Trim(paryCategories(enmfgID, i)) & "." & Trim(paryCategories(envendID, i)) & "'>" & pstrSubSubCatIndent & Server.HTMLEncode(Trim(paryCategories(envendName, i))) & "</option>" & vbcrlf
								pstrSubSubCatName = paryCategories(envendName, i)
							End If
						End If
						pstrSubCatName = paryCategories(enmfgName, i)
					Else
						If Len(pstrSubSubCatDBField) > 0 Then
							If paryCategories(envendName, i) <> pstrSubSubCatName Then
								If paryCategories(envendID, i) <> 1 Then Response.Write "      <option value='" & Trim(paryCategories(encatID, i)) & "." & Trim(paryCategories(enmfgID, i)) & "." & Trim(paryCategories(envendID, i)) & "'>" & pstrSubSubCatIndent & Server.HTMLEncode(Trim(paryCategories(envendName, i))) & "</option>" & vbcrlf
								pstrSubSubCatName = paryCategories(envendName, i)
							End If
						End If
					End If
				End If
			Next
			Response.Write "</select>"
		End If
		
	End Sub	'WriteSingleSelect

	'***********************************************************************************************

End Class	'ssCategorySE

'-----------------------------------------------------------

Sub WriteSearchFormSE

Dim pclsCategory

	Set pclsCategory = New ssCategorySE
	pclsCategory.LoadCategoryScript
	pclsCategory.WriteSearchForm
	Set pclsCategory = Nothing
	
End Sub	'WriteSearchFormSE

'-----------------------------------------------------------

Sub WriteCategoriesSE(strCurrentCategory, strCurrentSubCategory, strCurrentSubSubCategory)

Dim pclsCategory

	Set pclsCategory = New ssCategorySE
	With pclsCategory
		.CurrentCategory = strCurrentCategory
		.CurrentSubCategory = strCurrentSubCategory
		.CurrentSubSubCategory = strCurrentSubSubCategory
		.WriteCategories
	End With
	Set pclsCategory = Nothing
	
End Sub	'WriteCategoriesSE

'***********************************************************************************************

Sub writeCategoryTrailSE(ByVal strCurrentCategory, ByVal strCurrentSubCategory, byVal strCurrentSubSubCategory, byVal blnShowHome)

Dim pclsCategory

	If False Then
		Response.Write "<fieldset><legend>Sub writeCategoryTrailSE</legend>"
		Response.Write "txtsearchParamCat: " & txtsearchParamCat & "<br />"
		Response.Write "txtsearchParamMan: " & txtsearchParamMan & "<br />"
		Response.Write "txtsearchParamVen: " & txtsearchParamVen & "<br />"
		Response.Write "</fieldset>"
	End If
	
	Set pclsCategory = New ssCategorySE
	With pclsCategory
		.CurrentCategory = strCurrentCategory
		.CurrentSubCategory = strCurrentSubCategory
		.CurrentSubSubCategory = strCurrentSubSubCategory
		If Not blnShowHome Then .CategoryTrailShowHome = ""
		Response.Write .createTrail
	End With
	
	'Call pclsCategory.LoadManufacturersByCategoryID(strCurrentCategory)
	'Call pclsCategory.LoadCategoriesByManufacturerID(strCurrentSubCategory)
	
	Set pclsCategory = Nothing

End Sub	'writeCategoryTrailSE

'-----------------------------------------------------------

Sub WriteProductsSE

Dim pclsCategory

	Set pclsCategory = New ssCategorySE
	pclsCategory.WriteProducts
	Set pclsCategory = Nothing
	
End Sub	'WriteProductsSE

'-----------------------------------------------------------

Sub WriteSingleSelectSE

Dim pclsCategory

	Set pclsCategory = New ssCategorySE
	pclsCategory.WriteSingleSelect
	Set pclsCategory = Nothing
	
End Sub	'WriteSingleSelectSE

'-----------------------------------------------------------

%>
