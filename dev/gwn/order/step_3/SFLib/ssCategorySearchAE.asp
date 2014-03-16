<%
'********************************************************************************
'*   Category Search Tool for StoreFront 5.0 AE			                        *
'*   Release Version   2.00.001													*
'*   Release Date:     September 15, 2003										*
'*   Revision Date:    N/A														*
'*																				*
'*   Release Notes:                                                             *
'*																				*
'*   Release 2.00.002 (August 31, 2005)											*
'*	   - Added support for specified URLs										*
'*	   - Added support for href classes											*
'*																				*
'*   Release 2.00.001 (September 15, 2003)										*
'*	   - Initial Release														*
'*   Release Date      October 28, 2001											*
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'***********************************************************************************************

Class ssCategoryAE

Private cstrCatSpacer
Private cstrCatSymbol
Private cstrCatSymbolEnd
Private paryListImages
Private pstrSearchImage
Private pstrEmptyCellSpacer
Private pstrStartFont
Private pstrEndFont
Private cstrTrailSpacerText
Private pblnShowCurrentCategoryAsText
Private pstrCurrentCategoryTextWrapper
Private pblnUseShortLink

Private pstrBottomCategoryTargetPage
Private pstrIntermediateCategoryTargetPage
Private pblnDisplayCurrentCategory
Private pblnDisplayCurrentCategoryAsLink
Private pblnDisplayCurrentLevelOnly
Private pblnDisplayAllCategoriesAtCurrentLevel
Private pstrCategoryFilter
Private pstrSubcategoryFilter
Private plngNumCategories

Private pstrCacheName
Private plngCacheTime

'Internal working variables
Private paryCategories
Private plngCurrentUID
Private plngCurrentCategoryID
Private plngCurrentParentID
Private plngCurrentDepth
Private pstrCategoryTrail
Private pstrCategoryTrailShowHome
Private pstrCategoryTrailShowHomeURL
Private pstrCurrentCategoryName
Private pstrCurrentCategoryImage
Private pstrCurrentCategoryDescription
Private pstrCurrentCategoryURL

Private plngMaxDepth
Private plngStartingPosition
Private plngEndingPosition
Private plngStartingLevel
Private plngBackupPosition

Private pstrTDClass
Private pstrHRefClass
Private pstrHRefClass_SubCat
Private pstrHRefClass_SubSubCat
Private cstrProductLIStyle

Private pstrCategoryList
Private pclssfSearch

Private en_Trail_None
Private en_Trail_RootLevel
Private en_Trail_SameLevel
Private en_Trail_1stChild
Private en_Trail_InTrail

	'***********************************************************************************************

	Private Sub class_Terminate()
		On Error Resume Next
		Set pclssfSearch = Nothing
	End Sub
	Private Sub class_Initialize()
		
		If cblnDebugCategorySearchTool Then
			Session("ssDebug_CategorySearchToolAE") = "True"
			Response.Write "<table border=0><tr><th colspan=2>Category Search Tool for StoreFront 5.0 AE</th></tr>"
			Response.Write "<tr><td>Release Version: </td><td>2.00.001</td></tr>"
			Response.Write "<tr><td>Release Date: </td><td>September 15, 2003</td></tr>"
			Response.Write "<tr><td colspan=""2""><hr></td></tr>"
			Response.Write "<tr><th colspan=""2"">Debugging enabled</th></tr>"
			Response.Write "</table>"
		End If

	'////////////////////////////////////////////////////////////////////////////////
	'//
	'//		USER CONFIGURATION

		'Drop down configuration
		cstrCatSpacer = "  "
		cstrCatSymbol = "-"
		cstrCatSymbolEnd = ">"
		
		'Category Trail
		cstrTrailSpacerText = "&nbsp;>>&nbsp;"
		pstrCategoryTrailShowHome = "<span class=""categoryTrail"">Home</span>"
		pstrCategoryTrailShowHomeURL = "default.asp"
		pblnDisplayCurrentCategoryAsLink = False
		pblnShowCurrentCategoryAsText = True
		pstrCurrentCategoryTextWrapper = "<span class=""categoryTrailCurrentCategoryName"">{currentCategoryName}</span>"	'{currentCategoryName} must be retained
		
		'Quick Search configuration
		pstrSearchImage = "Search" 'a text example of a search button
		'pstrSearchImage = "<img src='" & C_BTN01 & "' alt=""Search"" border=""1"">" 'the default search button
		pblnDisplayCurrentLevelOnly = False
		pblnDisplayCurrentCategory = False
		
		'Category listing configuration
		Redim paryListImages(3)		'
		paryListImages(0) = ""
		paryListImages(1) = "-"
		paryListImages(2) = "--"
		paryListImages(3) = "---"
		
		paryListImages(0) = ""
		paryListImages(1) = "<img src=""images/arrow_left.gif"" alt=""-"" border=""0"">"
		paryListImages(2) = "<img src=""images/bullet_L2.gif"" alt=""-"" border=""0"">"
		paryListImages(3) = "<img src=""images/bullet_L3.gif"" alt=""-"" border=""0"">"
		
		'the following MUST be a complete td element
		pstrEmptyCellSpacer = "<td class=""{TDClass}""><img src=""images/transparent.gif"" alt=""-"" border=""0""></td>"
		
		pstrStartFont = "<span style=""vertical-align:middle"">" 'used to set font for category listing table
		pstrEndFont = "</span>"							'closing tag for category listing table font
		
		'Advanced settings
		pstrBottomCategoryTargetPage = "search_results.asp"
		pstrIntermediateCategoryTargetPage = "search_results.asp"
		pblnDisplayAllCategoriesAtCurrentLevel = False
		pblnUseShortLink  = False
		plngCacheTime = 600	'in seconds
		
		pstrTDClass = "clsMenuNavigationCat"
		pstrHRefClass = "clsMenuNavigationCat"
		pstrHRefClass_SubCat = "clsMenuNavigationCat"
		pstrHRefClass_SubSubCat = "clsMenuNavigationCat"
		cstrProductLIStyle = ""

	'//
	'//
	'////////////////////////////////////////////////////////////////////////////////
		
		en_Trail_None = -1
		en_Trail_RootLevel = 0
		en_Trail_SameLevel = 1
		en_Trail_1stChild = 2
		en_Trail_InTrail = 3

		Set pclssfSearch = New sfSearch
		pclssfSearch.LoadSavedSearchCriteria("searchCriteria")
		
	End Sub	'class_Initialize

	'***********************************************************************************************
	
	'Drop down configuration
	Public Property Let CatSpacer(byVal vntValue)
		cstrCatSpacer = vntValue
	End Property
	
	Public Property Let CatSymbol(byVal vntValue)
		cstrCatSymbol = vntValue
	End Property
	
	Public Property Let CatSymbolEnd(byVal vntValue)
		cstrCatSymbolEnd = vntValue
	End Property
	
	'Cache Info
	Public Property Let CacheName(byVal vntValue)
		pstrCacheName = vntValue
	End Property
	
	Public Property Let CacheTime(byVal vntValue)
		plngCacheTime = vntValue
	End Property

	Public Property Let HRefClass(byVal vntValue)
		pstrHRefClass = vntValue
	End Property

	Public Property Let HRefClass_SubCat(byVal vntValue)
		pstrHRefClass_SubCat = vntValue
	End Property

	Public Property Let HRefClass_SubSubCat(byVal vntValue)
		pstrHRefClass_SubSubCat = vntValue
	End Property

	Public Property Let TDClass(byVal vntValue)
		pstrTDClass = vntValue
	End Property
	
	'Category Trail
	Public Property Let TrailSpacerText(byVal vntValue)
		cstrTrailSpacerText = vntValue
	End Property
	
	Public Property Let CategoryTrailShowHome(byVal vntValue)
		pstrCategoryTrailShowHome = vntValue
	End Property
	
	Public Property Let CategoryTrailShowHomeURL(byVal vntValue)
		pstrCategoryTrailShowHomeURL = vntValue
	End Property
	
	Public Property Let ShowCurrentCategoryAsText(byVal vntValue)
		pblnShowCurrentCategoryAsText = vntValue
	End Property
	
	Public Property Let DisplayCurrentCategoryAsLink(byVal vntValue)
		pblnDisplayCurrentCategoryAsLink = vntValue
	End Property
	
	'Quick Search configuration
	Public Property Let SearchImage(byVal vntValue)
		pstrSearchImage = vntValue
	End Property
	
	Public Property Let DisplayAllCategoriesAtCurrentLevel(byVal vntValue)
		pblnDisplayAllCategoriesAtCurrentLevel = vntValue
	End Property
	
	Public Property Let DisplayCurrentCategory(byVal vntValue)
		pblnDisplayCurrentCategory = vntValue
	End Property
	
	Public Property Let IntermediateCategoryTargetPage(byVal vntValue)
		pstrIntermediateCategoryTargetPage = vntValue
	End Property
	
	Public Property Let BottomCategoryTargetPage(byVal vntValue)
		pstrBottomCategoryTargetPage = vntValue
	End Property
	
	Public Property Let CategoryFilter(byVal vntValue)
		pstrCategoryFilter = vntValue
	End Property
	
	Public Property Let SubcategoryFilter(byVal vntValue)
		pstrSubcategoryFilter = vntValue
	End Property
	
	Public Property Let DisplayCurrentLevelOnly(byVal vntValue)
		pblnDisplayCurrentLevelOnly = vntValue
	End Property
	
	Public Property Get CategoryTrail
		CategoryTrail = pstrCategoryTrail
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
	
	Public Property Get CurrentCategoryURL
		CurrentCategoryURL = pstrCurrentCategoryURL
	End Property
	
	Public Property Let UseShortLink(byVal vntValue)
		pblnUseShortLink = vntValue
	End Property
	
	Private Function getCategoryParent(ByVal strHier, ByVal lngCategoryID, ByVal lngDepth)
	
	Dim paryHier
	Dim plngCategoryID
	
		If lngDepth = 0 Then
			paryHier = Split(strHier, "-")
			If UBound(paryHier) >= (1) Then plngCategoryID = paryHier(1)
		ElseIf lngDepth = 1 Then
			plngCategoryID = strHier
		Else
			paryHier = Split(strHier, "-")
			If UBound(paryHier) >= (lngDepth-2) Then plngCategoryID = paryHier(lngDepth-2)
		End If
		
		getCategoryParent = plngCategoryID
	
	End Function	'getCategoryParent

	'***********************************************************************************************
	
	Public Function LoadCategories()
	
	Dim pblnLocalTest
	Dim prsCategory
	Dim pstrSQL
	Dim i
	
		If Err.number <> 0 Then Err.Clear
		
		'Only need to load if not previously loaded
		If isArray(paryCategories) Then
			LoadCategories = True
			Exit Function
		End If
	
		'Check Application for Values
		pblnLocalTest = False	'True	False	-For Testing
		If pblnLocalTest Then
			Response.Write "<h4>Application caching of categories disabled for testing</h4>"
			Call removeFromCache("ssCategorySearch" & pstrCacheName)
		End If

		If Not isCacheItemExpired("ssCategorySearch" & pstrCacheName) Then
			paryCategories = getFromCache("ssCategorySearch" & pstrCacheName)
			If cblnDebugCategorySearchTool Or pblnLocalTest Then Response.Write "<h4>Loading Categories from Application</h4>"
		End If
		
		If Not isArray(paryCategories) Then
			If cblnDebugCategorySearchTool Then Response.Write "<h4>Loading Categories from database</h4>"
			pstrSQL = "SELECT sfCategories.catID, sfSub_Categories.subcatID, sfCategories.catName, sfSub_Categories.subcatCategoryId, sfSub_Categories.CatHierarchy, sfSub_Categories.subcatName, sfSub_Categories.Depth, sfSub_Categories.subcatDescription, sfSub_Categories.subcatHttpAdd, sfSub_Categories.subcatImage, sfSub_Categories.subcatIsActive" _
					& " FROM sfSub_Categories RIGHT JOIN sfCategories ON sfSub_Categories.subcatCategoryId = sfCategories.catID" _
					& " Where catIsActive=1 AND sfSub_Categories.subcatIsActive=1" _
					& " ORDER BY sfSub_Categories.Depth, sfSub_Categories.subcatName"
'					& " Where catIsActive=1" _

			Set prsCategory = CreateObject("ADODB.RECORDSET")
			With prsCategory
				.CursorLocation = 2 'adUseClient
				If cblnDebugCategorySearchTool Then debugprint "pstrSQL", pstrSQL
				.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
				If cblnDebugCategorySearchTool Then debugprint "RecordCount", .RecordCount
				
				If Err.number <> 0 Then
					Response.Write "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />"
					Response.Write "Error in ssCategorySearch - LoadCategories: <br />SQL: " & sql & "<br />"
					.Close
					LoadCategories = False
					Response.Flush
					Exit Function
				Else
					On Error Goto 0
					Call LoadCategoryXML(prsCategory)

					'Save to Application
					If isArray(paryCategories) Then
						If cblnDebugCategorySearchTool Or pblnLocalTest Then Response.Write "<h4>Saving Categories to Application</h4>"
						Call saveToCache("ssCategorySearch" & pstrCacheName, paryCategories, DateAdd("s", plngCacheTime, Now()))
					End If

				End If

				.Close
			End With	'prsCategory
			set prsCategory = nothing
		End If
			
		'Now set the working variables
		If isArray(paryCategories) Then
			plngNumCategories = UBound(paryCategories)
		Else
			plngNumCategories = -1
		End If

		If Len(pstrSubcategoryFilter) > 0 Then
			plngCurrentCategoryID = pstrSubcategoryFilter
			plngCurrentUID = plngCurrentCategoryID
			For i = 0 To plngNumCategories
				If paryCategories(i, en_CatFields_uid) = CStr(plngCurrentCategoryID) Then
					plngCurrentDepth = paryCategories(i, en_CatFields_ParentLevel)
					plngCurrentParentID = paryCategories(i, en_CatFields_ParentID)
					pstrCurrentCategoryName = paryCategories(i, en_CatFields_Name)
					pstrCurrentCategoryImage = paryCategories(i, en_CatFields_ImagePath)
					pstrCurrentCategoryDescription = paryCategories(i, en_CatFields_Description)
					pstrCurrentCategoryURL = paryCategories(i, en_CatFields_URL)
					Exit For
				End If
			Next	'i
			Call setTrailStatuses(plngCurrentUID)
			pstrCategoryTrail = createTrail(plngCurrentUID)
		ElseIf Len(pstrCategoryFilter) > 0 Then
			plngCurrentCategoryID = pstrCategoryFilter
			plngCurrentUID = paryCategories(getArrayPositionByCategoryID(plngCurrentCategoryID), en_CatFields_UID)
			plngCurrentDepth = 0
			plngCurrentParentID = 0
			Call setTrailStatuses(plngCurrentUID)
			pstrCategoryTrail = createTrail(plngCurrentUID)
			pstrCurrentCategoryName = paryCategories(getArrayPositionByCategoryID(plngCurrentCategoryID), en_CatFields_Name)
			pstrCurrentCategoryImage = paryCategories(getArrayPositionByCategoryID(plngCurrentCategoryID), en_CatFields_ImagePath)
			pstrCurrentCategoryDescription = paryCategories(getArrayPositionByCategoryID(plngCurrentCategoryID), en_CatFields_Description)
			pstrCurrentCategoryURL = paryCategories(getArrayPositionByCategoryID(plngCurrentCategoryID), en_CatFields_URL)
		Else
			plngCurrentCategoryID = -1
			plngCurrentUID = -1
			plngCurrentDepth = -1
			plngCurrentParentID = -1
		End If

		If cblnDebugCategorySearchTool Then
			Response.Write "<b>Loading Categories</b><br />"
			Response.Write "&nbsp;&nbsp;Category Filter: " & pstrCategoryFilter & "<br />"
			Response.Write "&nbsp;&nbsp;Subcategory Filter: " & pstrSubcategoryFilter & "<br />"
			Response.Write "&nbsp;&nbsp;Current CategoryID: " & plngCurrentCategoryID & "<br />"
			Response.Write "&nbsp;&nbsp;Current UID: " & plngCurrentUID & "<br />"
			Response.Write "&nbsp;&nbsp;Current Depth: " & plngCurrentDepth & "<br />"
			Response.Flush
		End If
		
		LoadCategories = True
		
	End Function	'LoadCategories

	'***********************************************************************************************

	Private Sub LoadCategoryXML(byRef prsCategory)
	
	'Declare the values from the query
	Dim plnguid
	Dim plngParentLevel
	Dim pstrName
	Dim plngParentID
	Dim plngIsActive
	Dim pstrDescription
	Dim pstrURL
	Dim pstrImagePath
	
	Dim pstrHier
	Dim plngCategoryID
	Dim plngDepth

	Dim pobjXMLDoc 'As System.Xml.XmlElement
	Dim pobjXMLNode 'As System.Xml.XmlElement
	Dim pobjXMLCategoriesNode 'As System.Xml.XmlElement
	Dim pobjXMLCategoryNode 'As System.Xml.XmlElement
	Dim i, j

		With prsCategory
			Set pobjXMLDoc = CreateObject("MSXML.DOMDocument")
			Set pobjXMLCategoriesNode = pobjXMLDoc.CreateElement("Categories")
			pobjXMLDoc.AppendChild pobjXMLCategoriesNode
			Do While Not .EOF
				plnguid = Trim(.Fields("subcatID").Value & "")
				plngParentLevel = Trim(.Fields("Depth").Value & "")
				pstrName = Trim(.Fields("subcatName").Value & "")
				
				pstrHier = Trim(.Fields("CatHierarchy").Value & "")
				plngCategoryID = Trim(.Fields("subcatCategoryId").Value & "")
				plngDepth = Trim(.Fields("Depth").Value & "")
				plngParentID = getCategoryParent(pstrHier, plngCategoryID, plngDepth)
				plngIsActive = Trim(.Fields("subcatIsActive").Value & "")
				If Len(plngIsActive) = 0 Then plngIsActive = 0
				pstrDescription = Trim(.Fields("subcatDescription").Value & "")
				pstrURL = Trim(.Fields("subcatHttpAdd").Value & "")
				pstrImagePath = Trim(.Fields("subcatImage").Value & "")

				If cblnDebugCategorySearchTool Then
					Response.Write "<fieldset><legend>LoadCategoryXML - subcatID=" & plnguid & "</legend><font color=black>"
					Response.Write "pstrName: " & pstrName & "<br />"
					Response.Write "pstrHier: " & pstrHier & "<br />"
					Response.Write "plngParentID: " & plngParentID & "<br />"
					Response.Write "plngParentLevel: " & plngParentLevel & "<br />"
					Response.Write "plngCategoryID: " & plngCategoryID & "<br />"
					Response.Write "plngDepth: " & plngDepth & "<br />"
					Response.Write "</font></fieldset>"
				End If
				
				'all nodes are identified by their subCatID
				If plngParentLevel = 0 Then
					Set pobjXMLCategoryNode = CreateCategoryNode(plnguid, 0, pstrName, plngParentID, plngIsActive, pstrDescription, pstrURL, pstrImagePath, plngCategoryID, pobjXMLDoc)
					pobjXMLCategoriesNode.AppendChild pobjXMLCategoryNode
				Else
					'need to find the parent node for 1st level subCategories
					If Left(pstrHier, 5) = "none-" And plngParentLevel = 1 Then
						Set pobjXMLCategoryNode = CreateCategoryNode(plnguid, 0, pstrName, plngParentID, plngIsActive, pstrDescription, pstrURL, pstrImagePath, plngCategoryID, pobjXMLDoc)
						pobjXMLCategoriesNode.AppendChild pobjXMLCategoryNode
					Else
						If plngParentLevel = 1 Then plngParentID = FindTopCategoryNode(pobjXMLDoc, plngCategoryID)
						If cblnDebugCategorySearchTool Then Response.Write "<font color=red>plngParentID: " & plngParentID & "</font><br />"
						If Len(plngParentID) = 0 Then
							Set pobjXMLCategoryNode = CreateCategoryNode(plnguid, 0, pstrName, plngParentID, plngIsActive, pstrDescription, pstrURL, pstrImagePath, plngCategoryID, pobjXMLDoc)
							pobjXMLCategoriesNode.AppendChild pobjXMLCategoryNode
						Else
							Set pobjXMLCategoryNode = CreateCategoryNode(plnguid, plngDepth, pstrName, plngParentID, plngIsActive, pstrDescription, pstrURL, pstrImagePath, plngCategoryID, pobjXMLDoc)
							Set pobjXMLNode = GetXMLNodeByKey(pobjXMLDoc, plngParentID)
							If Not pobjXMLNode is Nothing Then
								pobjXMLNode.ChildNodes.Item(en_CatFields_IsBottom).Text = CStr(False)
								pobjXMLNode.AppendChild pobjXMLCategoryNode
							Else
								Response.Write "<h1>Unable to locate category node for <em>" & plngParentID & "</em></h1>"
							End If
						End If
					End If
				End If
				.MoveNext
			Loop
			'Response.Write "<hr>" & vbcrlf & vbcrlf & vbcrlf & pobjXMLDoc.XML & vbcrlf & vbcrlf & vbcrlf & "<hr>"
			
			'Now break out the categories
			Dim pobjNodeList 'As XmlNodeList
			Set pobjNodeList = pobjXMLDoc.GetElementsByTagName("Category")
			ReDim paryCategories(pobjNodeList.Length - 1, en_CatFields_NumFields)
			For i = 0 To pobjNodeList.Length - 1
				If pobjNodeList.Item(i).ChildNodes.Length > 0 Then
					For j = 0 To en_CatFields_NumFields
						paryCategories(i, j) = pobjNodeList.Item(i).ChildNodes.Item(j).Text
					Next 'j
				End If
			Next 'i

			If cblnDebugCategorySearchTool Then
				Response.Write "<fieldset><legend>LoadCategoryXML</legend><font color=red>" & Server.HTMLEncode(pobjXMLDoc.XML) & "</font><br />"
				
				Response.Write "<ul>"
				For i = 0 To pobjNodeList.Length - 1
					Response.Write "<li>" & j & ": " & paryCategories(i, 2) & "</li>"
					Response.Write "<ul>"
						Response.Write "<li>Description: " & paryCategories(i, 5) & "</li>"
						Response.Write "<li>URL: " & paryCategories(i, 10) & "</li>"
						Response.Write "<li>ImagePath: " & paryCategories(i, 6) & "</li>"
					Response.Write "</ul>"
				Next 'i
				Response.Write "</ul>"

				Response.Write "</fieldset>"
				Response.Flush
			End If

			If .RecordCount > 0 Then .MoveFirst
		End With	'prsCategory
		
		'Now Clean up
		Set pobjXMLDoc = Nothing
		Set pobjXMLCategoriesNode = Nothing
		Set pobjXMLCategoryNode = Nothing
		Set pobjXMLNode = Nothing
		Set pobjNodeList = Nothing
			
	End Sub	'LoadCategoryXML

	'***********************************************************************************************

	Private Function CreateCategoryNode(ByVal lnguid, ByVal lngParentLevel, ByVal strName, ByVal lngParentID, ByVal lngIsActive, ByVal strDescription, ByVal strURL, ByVal strImagePath, ByVal lngCategoryID, ByRef objXMLDoc)

		Dim pobjXMLCategoryNode 'As System.Xml.XmlElement
		Dim pobjXMLCategoryAttributeNode 'As System.Xml.XmlElement

		Set pobjXMLCategoryNode = objXMLDoc.CreateElement("Category")
		pobjXMLCategoryNode.SetAttribute "catID", lnguid

		'add the uid
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("uid")
		pobjXMLCategoryAttributeNode.Text = lnguid
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the ParentLevel
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("ParentLevel")
		pobjXMLCategoryAttributeNode.Text = lngParentLevel
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the name
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("Name")
		pobjXMLCategoryAttributeNode.Text = strName
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the ParentID
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("ParentID")
		pobjXMLCategoryAttributeNode.Text = lngParentID
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the IsActive
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("IsActive")
		pobjXMLCategoryAttributeNode.Text = lngIsActive
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the Description
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("Description")
		pobjXMLCategoryAttributeNode.Text = strDescription
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the URL
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("URL")
		pobjXMLCategoryAttributeNode.Text = strURL
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the ImagePath
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("ImagePath")
		pobjXMLCategoryAttributeNode.Text = strImagePath
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the CategoryID
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("CategoryID")
		pobjXMLCategoryAttributeNode.Text = lngCategoryID
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the IsBottom - Use default
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("IsBottom")
		pobjXMLCategoryAttributeNode.Text = "True"	'blnIsBottom
		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		'add the InTrail - Use default
		Set pobjXMLCategoryAttributeNode = objXMLDoc.CreateElement("InTrail")
		If lngParentLevel = 0 Then
			pobjXMLCategoryAttributeNode.Text = en_Trail_RootLevel
		Else
			pobjXMLCategoryAttributeNode.Text = en_Trail_None
		End If

		pobjXMLCategoryNode.AppendChild pobjXMLCategoryAttributeNode

		If cblnDebugCategorySearchTool Then
			Response.Write "<ul><li>Category " & lnguid & "</li>"
			Response.Write "<ul><li>Name: " & strName & "</li>"
			Response.Write "<li>ParentID: " & lngParentID & "</li>"
			Response.Write "<li>uid: " & lnguid & "</li>"
			Response.Write "<li>ParentLevel: " & lngParentLevel & "</li>"
			Response.Write "<li>IsActive: " & lngIsActive & "</li>"
			Response.Write "<li>Description: " & strDescription & "</li>"
			Response.Write "<li>URL: " & strURL & "</li>"
			Response.Write "<li>ImagePath: " & strImagePath & "</li>"
			Response.Write "<li>CategoryID: " & lngCategoryID & "</li>"
			Response.Write "</ul></ul>"
		End If
		Set CreateCategoryNode = pobjXMLCategoryNode

		'Now Clean up
		Set pobjXMLCategoryAttributeNode = Nothing
		Set pobjXMLCategoryNode = Nothing

	End Function 'CreateCategoryNode

	'***********************************************************************************************

	Function GetXMLNodeByKey(ByVal objXMLDoc, ByVal strKey)

	Dim nodeList
	Dim i
	
		Set nodeList = objXMLDoc.getElementsByTagName("Category")

		For i =0 To nodeList.length - 1
			If nodeList.Item(i).attributes.item(0).nodeValue = strKey Then
				Set GetXMLNodeByKey = nodeList.Item(i)
				Set nodeList = Nothing
				Exit Function
			End If
		Next 'i
		
		Set nodeList = Nothing
		Set GetXMLNodeByKey = Nothing

	End Function 'GetXMLNodeByKey

	'***********************************************************************************************

	Function FindTopCategoryNode(ByVal objXMLDoc, ByVal lngCategoryID)
	'Find the node with the specified categoryID and ParentLevel = 0
	'Return: uid
	
	Dim nodeList
	Dim i
	Dim pstrID
	Dim pblnLocalDebug:	pblnLocalDebug = False
	Dim pblnMatch
	
		Set nodeList = objXMLDoc.getElementsByTagName("Category")

		'If pblnLocalDebug Then Response.Write vbcrlf & vbcrlf & "<br />FindTopCategoryNode: (" & objXMLDoc.xml & ")<br />" & vbcrlf & vbcrlf
		If pblnLocalDebug Then Response.Write "<ul><li>FindTopCategoryNode: Node Count =" & nodeList.length & " - Looking for <em>" & lngCategoryID & "</em></li><ul>"
		For i = 0 To nodeList.length - 1
			If nodeList.Item(i).ChildNodes.Item(en_CatFields_ParentLevel).Text = "0" Then
				pblnMatch = CBool(nodeList.Item(i).ChildNodes.Item(en_CatFields_CategoryID).Text = CStr(lngCategoryID))
				If pblnLocalDebug Then Response.Write "<li>Node " & i & " - Match: " & pblnMatch & " (<i>UID: " & nodeList.Item(i).ChildNodes.Item(en_CatFields_uid).Text & "</i>) - CatID: " & nodeList.Item(i).ChildNodes.Item(en_CatFields_CategoryID).Text & "</li>"
				If pblnMatch Then
					pstrID = nodeList.Item(i).ChildNodes.Item(en_CatFields_uid).Text
					Exit For
				End If
			End If
		Next 'i
		If pblnLocalDebug Then
			Response.Write "<li>FindTopCategoryNode: (" & nodeList.length & ") - " & lngCategoryID & ": " & pstrID & "</li>"
			Response.Write "</ul></ul>"
		End If
		
		Set nodeList = Nothing
		
		FindTopCategoryNode = pstrID

	End Function 'FindTopCategoryNode

	'***********************************************************************************************

	Public Function SelectCategoryAE(byVal strSelectedSubCatID, byVal bytSize)

	Dim pstrCatName
	Dim pstrCatValue
	Dim pstrsubCatName
	Dim pstrsubCatValue
	Dim i
	Dim pstrTemp
	
		Call LoadCategories
		If Len(strSelectedSubCatID) > 0 And isNumeric(strSelectedSubCatID) Then plngCurrentUID = strSelectedSubCatID
		pstrTemp = "<select id='CatSource' name='CatSource' size='" & bytSize & "' onchange='var theForm=this.form; var pstrTemp=this.value.split(""."");theForm.txtsearchParamCat.value=pstrTemp[0];theForm.subcat.value=pstrTemp[1];theForm.iLevel.value=pstrTemp[2];theForm.txtCatName.value=pstrTemp[2];'>" & vbcrlf _
					& WriteOption("ALL..1","All Categories",False)
		pstrTemp = pstrTemp & getDropdownData & vbcrlf
		pstrTemp = pstrTemp & "</select>" & vbcrlf
			
		SelectCategoryAE = pstrTemp
		
	End Function	'SelectCategoryAE

	'***********************************************************************************************

	Public Function getDropdownData

	Dim pstrCatName
	Dim pstrCatValue
	Dim pstrsubCatName
	Dim pstrsubCatValue
	Dim i
	Dim pstrTemp
	Dim pblnSelected
	
		Call LoadCategories

		For i = 0 To plngNumCategories

			If paryCategories(i, en_CatFields_ParentLevel) = 0 Then
				pstrCatValue = paryCategories(i, en_CatFields_uid) & "..1"
				pstrCatValue = paryCategories(i, en_CatFields_CategoryID) & "..1"
				pblnSelected = CBool(Trim(plngCurrentUID) =  Trim(paryCategories(i, en_CatFields_CategoryID)))
				pstrCatName = paryCategories(i, en_CatFields_Name)
			ElseIf paryCategories(i, en_CatFields_ParentLevel) > 0 Then
				pstrCatValue = paryCategories(i, en_CatFields_uid) & "." & paryCategories(i, en_CatFields_uid) & "." & cStr(paryCategories(i, en_CatFields_ParentLevel)+1)
				pstrCatValue = paryCategories(i, en_CatFields_CategoryID) & "." & paryCategories(i, en_CatFields_uid) & "." & cStr(paryCategories(i, en_CatFields_ParentLevel)+1)
				pstrCatName = String(paryCategories(i, en_CatFields_ParentLevel),cstrCatSymbol) & cstrCatSymbolEnd & cstrCatSpacer & paryCategories(i, en_CatFields_Name)
				pblnSelected = CBool(Trim(plngCurrentUID) =  Trim(paryCategories(i, en_CatFields_uid)))
			End If
			pstrTemp = pstrTemp & WriteOption(pstrCatValue, pstrCatName, pblnSelected)
		Next
			
		getDropdownData = pstrTemp
		
	End Function	'getDropdownData

	'***********************************************************************************************
	
	Public Sub StoreSearchSettings

		With pclssfSearch
			.iLevel = iLevel
			.subcat = sSubCat
			.txtFromSearch = txtFromSearch
			.txtsearchParamTxt = txtsearchParamTxt
			.txtsearchParamType = txtsearchParamType
			.txtsearchParamCat = txtsearchParamCat
			.txtsearchParamMan = txtsearchParamMan
			.txtsearchParamVen = txtsearchParamVen
			.txtDateAddedStart = txtDateAddedStart
			.txtDateAddedEnd = txtDateAddedEnd
			.txtPriceStart = txtPriceStart
			.txtPriceEnd = txtPriceEnd
			.txtSale = txtSale
			.SaveSearchCriteria("searchCriteria")
			If cblnDebugCategorySearchTool Then
				Response.Write "<h4>Saving search criteria to cache . . .</h4>"
				.WriteSearchCriteria
			End If
		End With
		
	End Sub	'StoreSearchSettings

	Public Sub LoadSearchSettings
	
		On Error Resume Next
	
		With pclssfSearch
			.LoadSavedSearchCriteria("searchCriteria")
			If cblnDebugCategorySearchTool Then
				Response.Write "<h4>Loading search criteria from cache . . .</h4>"
				.WriteSearchCriteria
			End If
			iLevel = .iLevel
			sSubCat = .subcat
			txtFromSearch = .txtFromSearch
			txtsearchParamTxt = .txtsearchParamTxt
			txtsearchParamType = .txtsearchParamType
			txtsearchParamCat = .txtsearchParamCat
			txtsearchParamMan = .txtsearchParamMan
			txtsearchParamVen = .txtsearchParamVen
			txtDateAddedStart = .txtDateAddedStart
			txtDateAddedEnd = .txtDateAddedEnd
			txtPriceStart = .txtPriceStart
			txtPriceEnd = .txtPriceEnd
			txtSale = .txtSale
			txtCatName = .txtCatName
		End With
		
		If Err.number <> 0 Then Err.Clear
		
	End Sub	'LoadSearchSettings

	'***********************************************************************************************
	
	Public Sub WriteSearchForm(bytSize)

		With Response
			.Write "<form id='frmCategory' name='frmCategory' action='" & pstrBottomCategoryTargetPage & "' method='Get'>" & vbcrlf
'Response.Write WriteHiddenField("txtsearchParamTxt","")
			.Write WriteHiddenField("txtFromSearch","fromSearch")
			.Write WriteHiddenField("txtsearchParamType","ALL")
			.Write WriteHiddenField("txtsearchParamMan","ALL")
			.Write WriteHiddenField("txtsearchParamVen","ALL")

			.Write WriteHiddenField("txtsearchParamCat","ALL")	'this is the CatID
			.Write WriteHiddenField("subcat","")	'this is the subCatID
			.Write WriteHiddenField("iLevel","1")	'this is the Depth
			.Write WriteHiddenField("txtCatName","") 'this is the Depth for whatever reason ??
			.Write "<table border=""0"" cellspacing=""0"" cellpadding=""3"">" & vbcrlf
			.Write "  <tr>" & vbcrlf
			.Write "    <td>"
			.Write "<b>Search</b>&nbsp;<input name='txtsearchParamTxt' value=''/>"
			.Write "&nbsp;<b>In</b>&nbsp;"
			.Write SelectCategoryAE("", bytSize)
			.Write "&nbsp;<a href='' onclick='document.frmCategory.submit(); return false;'>" & pstrSearchImage & "</a></td>" & vbcrlf
			.Write "    </td>" & vbcrlf
			.Write "  </tr>" & vbcrlf
			.Write "  <tr>" & vbcrlf
			.Write "    <td>" & vbcrlf
			.Write "<input type='radio' value='ALL' checked name='txtsearchParamType'>&nbsp;<b>ALL</b>&nbsp;Words&nbsp;<input type='radio' name='txtsearchParamType' value='ANY'>&nbsp;<b>ANY</b>&nbsp;Words&nbsp;<input type='radio' name='txtsearchParamType' value='Exact'>&nbsp;Exact&nbsp;Phrase" & vbcrlf
			.Write "    </td>" & vbcrlf
			.Write "  </tr>" & vbcrlf
			.Write "</table>" & vbcrlf
			.Write "</form>" & vbcrlf
		End With

	End Sub	'WriteSearchForm

	'***********************************************************************************************

	Private Sub SetStartupPositions()
	
	Dim i
	
		'Determine max depth of categories
		plngMaxDepth = 0
		For i = 0 To plngNumCategories
			If paryCategories(i, en_CatFields_ParentLevel) > plngMaxDepth Then plngMaxDepth = paryCategories(i, en_CatFields_ParentLevel)
		Next	'i
		plngMaxDepth = plngMaxDepth + 1
		
		'Determine items to display
		'pblnDisplayCurrentLevelOnly = False
		'debugprint "plngMaxDepth", 	plngMaxDepth	
		'Response.Write "<h1>plngCurrentCategoryID: " & plngCurrentCategoryID & "</h1>"	
		If pblnDisplayCurrentLevelOnly Then
			If plngCurrentDepth = -1 Then
				plngStartingPosition = 0
				plngStartingLevel = 0
			ElseIf plngCurrentDepth = 0 Then
				plngStartingPosition = getArrayPositionByCategoryID(plngCurrentCategoryID)
				plngStartingLevel = paryCategories(plngStartingPosition, en_CatFields_ParentLevel)
			ElseIf plngCurrentCategoryID = 0 Then
				plngStartingPosition = getArrayPositionByCategoryID(plngCurrentCategoryID)
				plngStartingLevel = paryCategories(plngStartingPosition, en_CatFields_ParentLevel)
			Else
				plngStartingPosition = getArrayPositionByID(plngCurrentCategoryID)
				plngStartingLevel = paryCategories(plngStartingPosition, en_CatFields_ParentLevel)
			End If
			plngBackupPosition = plngStartingLevel
			
			plngEndingPosition = plngStartingPosition
			For i = plngStartingPosition To 0 Step -1
				If paryCategories(i, en_CatFields_ParentLevel) < plngStartingLevel Then Exit For
				plngBackupPosition = i
			Next 'i
			plngStartingPosition = plngBackupPosition

		Else
			plngStartingPosition = 0
			plngStartingLevel = 0
		End If
		
		If cblnDebugCategorySearchTool Then
			Response.Write "<fieldset style='color:blue;'><legend>Setting Startup positions</legend><font background=white>"
			Response.Write "&nbsp;&nbsp;plngNumCategories: " & plngNumCategories & "<br />"
			Response.Write "&nbsp;&nbsp;pblnDisplayCurrentLevelOnly: " & pblnDisplayCurrentLevelOnly & "<br />"
			Response.Write "&nbsp;&nbsp;plngMaxDepth: " & plngMaxDepth & "<br />"
			Response.Write "&nbsp;&nbsp;CurrentCategoryID: " & plngCurrentCategoryID & "<br />"
			Response.Write "&nbsp;&nbsp;plngStartingPosition: " & plngStartingPosition & "<br />"
			Response.Write "&nbsp;&nbsp;plngStartingLevel: " & plngStartingLevel & "<br />"
			Response.Write "&nbsp;&nbsp;plngBackupPosition: " & plngBackupPosition & "<br />"
			Response.Write "</font></fieldset>"
		End If

	End Sub	'SetStartupPositions

	'***********************************************************************************************

	Public Sub WriteCategoriesAsLinks

	Dim i
	Dim pstrTemp
	Dim pstrLineOut
	Dim pblnDisplay
	Dim plngColumns
	
		Call LoadCategories
		Call SetStartupPositions

		pstrEmptyCellSpacer = "<img src=""images/transparent.gif"" alt=""-"" border=""0"" class=""imgNavButtsIndent"">"
		pstrStartFont = ""
		pstrEndFont = ""
		pstrHRefClass = ""
		pstrHRefClass_SubCat = "aNavButtsIndent1"
		pstrHRefClass_SubSubCat = "aNavButtsIndent2"

		pstrTemp = "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" id=""navButts"">" _
				 & "<tr><td>"

		plngColumns = plngCurrentDepth + 2	'0 based array + 1 for potential subcategory

		For i = 0 To plngNumCategories
			pstrLineOut = ""
			
			If paryCategories(i, en_CatFields_CategoryID) <> 1 Then	'Only display categories other than "No Category"
				If pblnDisplayCurrentLevelOnly Then
					Select Case CLng(paryCategories(i, en_CatFields_InTrail))
						Case en_Trail_None
							pblnDisplay = False
						Case en_Trail_RootLevel
							pblnDisplay = True
						Case en_Trail_SameLevel
							pblnDisplay = True
						Case en_Trail_1stChild
							pblnDisplay = True
						Case en_Trail_InTrail
							pblnDisplay = True
						Case Else
							pblnDisplay = True
					End Select
					
					If pblnDisplay Then
						pstrLineOut = Repeat(paryCategories(i, en_CatFields_ParentLevel), pstrEmptyCellSpacer) _
									& GetCategoryImage(i) & CategoryLink(i, False) & vbcrlf
					End If
					
				Else		
					If CBool(CLng(plngCurrentDepth) = CLng(paryCategories(i, en_CatFields_ParentLevel))) And CBool(plngCurrentUID = paryCategories(i, en_CatFields_uid)) Then
						If pblnDisplayCurrentCategory Then pstrLineOut = Repeat(paryCategories(i, en_CatFields_ParentLevel), pstrEmptyCellSpacer) & GetCategoryImage(i) & CategoryLink(i, False) & vbcrlf
					Else
						pstrLineOut = Repeat(paryCategories(i, en_CatFields_ParentLevel), pstrEmptyCellSpacer) & GetCategoryImage(i) & CategoryLink(i, False) & vbcrlf
					End if
				End If	'pblnDisplayCurrentLevelOnly
				
				pstrTemp = pstrTemp & pstrLineOut
				
			End If	'paryCategories(i, en_CatFields_CategoryID) <> 1
			
		Next	'i
		pstrTemp = pstrTemp & "</td></tr></table>" & vbcrlf
		
		Response.Write pstrTemp
			
	End Sub	'WriteCategoriesAsLinks

	'***********************************************************************************************

	Public Sub WriteCategoriesAsTable

	Dim i
	Dim pstrTemp
	Dim pstrLineOut
	Dim pblnDisplay
	Dim plngColumns
	Dim pstrTempDebug
	
		Call DebugRecordSplitTime("LoadCategories . . .(SFLib/ssCategorySearchAE.asp?WriteCategoriesAsTable)")
		Call LoadCategories
		Call DebugRecordSplitTime("LoadCategories complete, beginning write . . .")
		Call SetStartupPositions

		pstrTemp = pstrTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf
		'pstrTemp = pstrTemp & "<tr><td colspan=""" & plngMaxDepth & """ class=""" & pstrTDClass & """ valign=""middle"">" & pstrCategoryTrail & "</td>" & vbcrlf

		plngColumns = plngCurrentDepth + 2	'0 based array + 1 for potential subcategory

		'pstrTempDebug = "<fieldset><legend>WriteCategoriesAsTable</legend>"
		'pstrTempDebug = pstrTempDebug & "plngColumns: " & plngColumns & "<br />" & vbcrlf
		'pstrTempDebug = pstrTempDebug & "plngMaxDepth: " & plngMaxDepth & "<br />" & vbcrlf

		For i = 0 To plngNumCategories
			pstrLineOut = ""
			
			If paryCategories(i, en_CatFields_CategoryID) <> 1 Then	'Only display categories other than "No Category"
				If pblnDisplayCurrentLevelOnly Then
					Select Case CLng(paryCategories(i, en_CatFields_InTrail))
						Case en_Trail_None
							pblnDisplay = False
						Case en_Trail_RootLevel
							pblnDisplay = True
						Case en_Trail_SameLevel
							pblnDisplay = True
						Case en_Trail_1stChild
							pblnDisplay = True
						Case en_Trail_InTrail
							pblnDisplay = True
						Case Else
							pblnDisplay = True
					End Select
					
					If cblnDebugCategorySearchTool Then	'Or True 
						Response.Write "<fieldset style='background-color:white; color=green; font-size:8pt;'><legend>WriteCategoriesAsTable (<b>" & i & "</b>)</legend>"
						Response.Write "&nbsp;&nbsp;Name: " & paryCategories(i, en_CatFields_Name) & "<br />"
						Response.Write "&nbsp;&nbsp;uid: " & paryCategories(i, en_CatFields_uid) & "<br />"
						Response.Write "&nbsp;&nbsp;ParentID: " & paryCategories(i, en_CatFields_ParentID) & "<br />"
						Response.Write "&nbsp;&nbsp;ParentLevel: " & paryCategories(i, en_CatFields_ParentLevel) & ""
						Response.Write "&nbsp;&nbsp;InTrail: " & paryCategories(i, en_CatFields_InTrail) & ""
						Response.Write "&nbsp;&nbsp;pblnDisplay: " & pblnDisplay & ""
						Response.Write "<hr />"
						Response.Write "&nbsp;&nbsp;plngColumns: " & plngColumns & "<br />"
						Response.Write "&nbsp;&nbsp;plngCurrentUID: " & plngCurrentUID & "<br />"
						Response.Write "&nbsp;&nbsp;plngCurrentDepth: " & plngCurrentDepth & "<br />"
						Response.Write "</fieldset>"
					End If

					If pblnDisplay Then
						pstrLineOut = "<tr>" _
									& Repeat(paryCategories(i, en_CatFields_ParentLevel), Replace(pstrEmptyCellSpacer, "{TDClass}", pstrTDClass)) _
									& "<td colspan=""" & cStr(plngColumns - paryCategories(i, en_CatFields_ParentLevel)) & """ class=""" & pstrTDClass & """>" _
									& GetCategoryImage(i) & CategoryLink(i, False) _
									& "</td>" _
									& "</tr>" & vbcrlf
					End If
					
				Else		
					'pstrTempDebug = pstrTempDebug & "Category:&nbsp;" & Replace(paryCategories(i, en_CatFields_Name), " ", "&nbsp;") & "&nbsp;-&nbsp;Depth:&nbsp;" & plngCurrentDepth & "<br />" & vbcrlf
					If CBool(CLng(plngCurrentDepth) = CLng(paryCategories(i, en_CatFields_ParentLevel))) And CBool(plngCurrentUID = paryCategories(i, en_CatFields_uid)) Then
						'If pblnDisplayCurrentCategory Then pstrLineOut = "<tr>" & Repeat(paryCategories(i, en_CatFields_ParentLevel)-plngCurrentDepth,"<td class=""" & pstrTDClass & """>&nbsp;</td>") & "<td colspan=""" & cStr(plngMaxDepth - paryCategories(i, en_CatFields_ParentLevel)) & """ class=""" & pstrTDClass & """>" & GetCategoryImage(i) & pstrStartFont & CategoryLink(i, False) & pstrEndFont & "</td></tr>" & vbcrlf
						If pblnDisplayCurrentCategory Then pstrLineOut = "<tr>" & Repeat(paryCategories(i, en_CatFields_ParentLevel),"<td class=""" & pstrTDClass & """>&nbsp;</td>") & "<td colspan=""" & cStr(plngMaxDepth - paryCategories(i, en_CatFields_ParentLevel)) & """ class=""" & pstrTDClass & """>" & GetCategoryImage(i) & CategoryLink(i, False) & "</td></tr>" & vbcrlf
					Else
						pstrLineOut = "<tr>" & Repeat(paryCategories(i, en_CatFields_ParentLevel),"<td class=""" & pstrTDClass & """>&nbsp;</td>") & "<td colspan=""" & cStr(plngMaxDepth - paryCategories(i, en_CatFields_ParentLevel)) & """ class=""" & pstrTDClass & """>" & GetCategoryImage(i) & CategoryLink(i, False) & "</td></tr>" & vbcrlf
					End if
				End If	'pblnDisplayCurrentLevelOnly
				
				pstrTemp = pstrTemp & pstrLineOut
				
			End If	'paryCategories(i, en_CatFields_CategoryID) <> 1
			
		Next	'i
		pstrTemp = pstrTemp & "</table>" & vbcrlf
		
		Response.Write pstrTemp
		'Response.Write pstrTempDebug
		Call DebugRecordSplitTime("LoadCategories complete")
			
	End Sub	'WriteCategoriesAsTable

	'***********************************************************************************************

	Public Sub WriteCategoriesAsList

	Dim i,j
	Dim pstrTemp
	Dim pblnWriteImmediately
	Dim pstrLineOut
	Dim pblnDisplay
	Dim pstrPrevDepth
	
		pblnWriteImmediately = False
		Call LoadCategories
		Call SetStartupPositions
		
		'If pblnDisplayCurrentLevelOnly Then plngMaxDepth = 2

		pstrPrevDepth = CLng(plngCurrentDepth)
		If pstrPrevDepth < 0 Then pstrPrevDepth = 0
		
		If cblnDebugCategorySearchTool Then
			debugprint "pstrPrevDepth", pstrPrevDepth
			debugprint "plngStartingPosition", plngStartingPosition
			debugprint "en_CatFields_ParentLevel", paryCategories(plngStartingPosition, en_CatFields_ParentLevel)
		End If

		pstrTemp = "<ul>" & vbcrlf

		pblnDisplay = True
		For i = plngStartingPosition To plngNumCategories
			pstrLineOut = ""
			
			If CLng(paryCategories(i, en_CatFields_ParentLevel)) > CLng(pstrPrevDepth) Then
				pstrTemp = pstrTemp & "<ul>" & vbcrlf
			ElseIf CLng(paryCategories(i, en_CatFields_ParentLevel)) < pstrPrevDepth Then
				For j = paryCategories(i, en_CatFields_ParentLevel) To pstrPrevDepth-1
					pstrTemp = pstrTemp & "</ul>" & vbcrlf
				Next 'j			
			End If
			
			If paryCategories(i, en_CatFields_CategoryID) <> 1 Then
				If pblnDisplayCurrentLevelOnly Then
					If CBool(CLng(plngCurrentDepth) = CLng(paryCategories(i, en_CatFields_ParentLevel))) Then
						If CBool(plngCurrentUID = paryCategories(i, en_CatFields_uid)) Then
							If pblnDisplayCurrentCategory Then pstrLineOut = "<li>" & pstrStartFont & Server.HTMLEncode(paryCategories(i, en_CatFields_Name)) & pstrEndFont & "</li>" & vbcrlf
							pblnDisplay = True
						ElseIf CBool(CLng(plngCurrentParentID) = CLng(paryCategories(i, en_CatFields_ParentID))) Or CBool(plngCurrentDepth = 0) Then
							If pblnDisplayAllCategoriesAtCurrentLevel Then pstrLineOut = "<li>" & pstrStartFont & Server.HTMLEncode(paryCategories(i, en_CatFields_Name)) & pstrEndFont & "</li>" & vbcrlf
							pblnDisplay = False
						End If
					ElseIf CLng(plngCurrentDepth) = CLng((paryCategories(i, en_CatFields_ParentLevel)-1)) Then
						If pblnDisplay Then pstrLineOut = "<li>" & CategoryLink(i, False) & "</li>" & vbcrlf
					ElseIf CLng(plngCurrentDepth) > CLng(paryCategories(i, en_CatFields_ParentLevel)-1) Then
						pblnDisplay = False
					End If
				Else		
					If CBool(CLng(plngCurrentDepth) = CLng(paryCategories(i, en_CatFields_ParentLevel))) And CBool(plngCurrentUID = paryCategories(i, en_CatFields_uid)) Then
						If pblnDisplayCurrentCategory Then pstrLineOut = "<li>" & pstrStartFont & Server.HTMLEncode(paryCategories(i, en_CatFields_Name)) & pstrEndFont & "</li>" & vbcrlf
					Else
						pstrLineOut = "<li>" & CategoryLink(i, False) & "</li>" & vbcrlf
					End if
				End If
				pstrTemp = pstrTemp & pstrLineOut
			End If
			
			If pblnWriteImmediately Then
				Response.Write pstrTemp
				pstrTemp = ""
			End If
			
			'This was here to experiment with a product listing by category
			'Dim pstrSQL
			'Function getProductSQLAE(searchParamType, searchParamTxt, searchParamCat, searchParamMan, searchParamVen, DateAddedStart, DateAddedEnd, PriceStart, PriceEnd, Sale, subCatID, Ilevel)
			'If paryCategories(i, en_CatFields_ParentLevel) = 0 Then
			'	pstrSQL = getProductSQLAE("fromSearch", "", cStr(paryCategories(i, en_CatFields_CategoryID)), "", "ALL", "ALL", "", "", "", "", cStr(paryCategories(i, en_CatFields_CategoryID)), 1)
			'Else
			'	pstrSQL = getProductSQLAE("fromSearch", "", cStr(paryCategories(i, en_CatFields_CategoryID)), "", "ALL", "ALL", "", "", "", "", paryCategories(i, en_CatFields_uid), cStr(paryCategories(i, en_CatFields_ParentLevel)+1))
			'End If

			'Response.Write "<li>" & paryCategories(i, en_CatFields_uid) & "</li>"
			'Response.Write "<li>" & paryCategories(i, en_CatFields_CategoryID) & "</li>"
			'Response.Write "<li>" & cStr(paryCategories(i, en_CatFields_ParentLevel)+1) & "</li>"
			'Response.Write "<li>" & pstrSQL & "</li>"
			'Call WriteProducts(pstrSQL)

			pstrPrevDepth = CLng(paryCategories(i, en_CatFields_ParentLevel))
		Next	'i
		pstrTemp = pstrTemp & "</ul>" & vbcrlf
		
		Response.Write pstrTemp
			
	End Sub	'WriteCategoriesAsList

	'***********************************************************************************************

	Public Sub WriteProducts(byVal strSQL, byVal blnShowCategoryTrail)

	Dim pstrSQL
	Dim prsProducts
	Dim pstrProdLink
	Dim i

		If Len(strSQL) > 0 Then
			pstrSQL = strSQL
		Else
			pstrSQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodShortDescription, sfProducts.prodLink" _
					& " FROM sfProducts" _
					& " WHERE sfProducts.prodEnabledIsActive=1" _
					& " ORDER BY sfProducts.prodName"
		End If

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

		Response.Write "<ul>" & vbcrlf
		With prsProducts
			
			For i = 1 to .RecordCount
				
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
					pstrProdLink = "<a " & pstrHRefClass & " href=" & Chr(34) & pstrProdLink & Chr(34) & " title=" & Chr(34) & .Fields("prodName").Value & Chr(34) & ">" & .Fields("prodName").Value & "</a>"
				Else
					If Len(.Fields("prodLink").Value & "") > 0 Then
						pstrProdLink = "<a " & pstrHRefClass & " href=" & Chr(34) & pstrProdLink & Chr(34) & "><img border=0 src=" & Chr(34) & .Fields("prodImageSmallPath").Value & Chr(34) & ">&nbsp;"& .Fields("prodName").Value & "</a>"
					Else
						pstrProdLink = "<a " & pstrHRefClass & " href=" & Chr(34) & pstrProdLink & Chr(34) & ">" & .Fields("prodName").Value & "</a>"
					End If
				End If
				Response.Write "<li " & cstrProductLIStyle & ">" & pstrProdLink
				If blnShowCategoryTrail Then
					Response.Write " | "
					Call writeDetailCategoryTrailAE(.Fields("prodID").Value, False)
				End If	'blnShowCategoryTrail
				Response.Write "</li>" & vbcrlf
				.MoveNext
			Next
			
			.Close
		End with
		Set prsProducts = Nothing
		
		Response.Write "</ul>" & vbcrlf

	End Sub	'WriteProducts

	'***********************************************************************************************
	
	Private Function GetCategoryImage(byVal lngCounter)
	
	Dim pstrTempImage
	
		If paryCategories(lngCounter, en_CatFields_ParentLevel) >= 0 Then
			If CLng(paryCategories(lngCounter, en_CatFields_ParentLevel)) > UBound(paryListImages) Then
				pstrTempImage = paryListImages(UBound(paryListImages))
			Else
				pstrTempImage = paryListImages(paryCategories(lngCounter, en_CatFields_ParentLevel))
			End If
		Else
			pstrTempImage = ""
		End If
		
		GetCategoryImage = pstrTempImage

	End Function	'GetCategoryImage
	
	'***********************************************************************************************
	
	Private Function getCustomURL(byVal strTempTarget, byVal blnCategory, byVal strCustomURL)
	
		If Len(strCustomURL) = 0 Then
			getCustomURL = strTempTarget & "?" & pclssfSearch.SearchLinkParameters(blnCategory)
		Else
			getCustomURL = strCustomURL
		End If

	End Function	'getCustomURL
	
	'***********************************************************************************************
	
	Private Function CategoryLink(byVal lngCounter, byVal blnCategory)
	'Input: lngCounter - array pointer
	'	  : blnCategory - True returns category link; False returns subCategory link
	'Output: link
	
	Dim pstrLink
	Dim pstrTempTarget
	Dim pstrTempURL
	Dim pstrClass
	
		If CBool(paryCategories(lngCounter, en_CatFields_IsBottom)) Then
			pstrTempTarget = pstrBottomCategoryTargetPage
			If Len(pstrTempTarget) = 0 Then pstrTempTarget = pstrIntermediateCategoryTargetPage
		Else
			pstrTempTarget = pstrIntermediateCategoryTargetPage
			If Len(pstrTempTarget) = 0 Then pstrTempTarget = pstrBottomCategoryTargetPage
		End If
		
		With pclssfSearch
			.txtsearchParamCat  = paryCategories(lngCounter, en_CatFields_CategoryID)
			.UseShortLink = pblnUseShortLink
			If blnCategory OR paryCategories(lngCounter, en_CatFields_ParentLevel) = 0 Then
				pstrTempURL = getCustomURL(pstrTempTarget, True, paryCategories(lngCounter, en_CatFields_URL))
				pstrClass = pstrHRefClass
			Else
				.subcat = paryCategories(lngCounter, en_CatFields_uid)
				.iLevel  = cStr(paryCategories(lngCounter, en_CatFields_ParentLevel)+1)
				.txtCatName = cStr(paryCategories(lngCounter, en_CatFields_ParentLevel)+1)
				pstrTempURL = getCustomURL(pstrTempTarget, blnCategory, paryCategories(lngCounter, en_CatFields_URL))
				
				'pstrTempURL = pstrTempTarget & "?txtsearchParamType=ALL&txtsearchParamMan=ALL&txtsearchParamVen=ALL&txtsearchParamTxt=*&txtFromSearch=fromSearch&txtsearchParamCat=" & paryCategories(lngCounter, en_CatFields_CategoryID) & "&subcat=" & paryCategories(lngCounter, en_CatFields_uid) & "&iLevel=" & cStr(paryCategories(lngCounter, en_CatFields_ParentLevel)+1) & "&txtCatName=" & cStr(paryCategories(lngCounter, en_CatFields_ParentLevel)+1)

				If paryCategories(lngCounter, en_CatFields_ParentLevel) = 1 Then
					pstrClass = pstrHRefClass_SubCat
				Else
					pstrClass = pstrHRefClass_SubSubCat
				End If
			End If

			If Len(pstrClass) > 0 Then
				pstrLink = "<a href=" & Chr(34) & pstrTempURL & Chr(34) _
						 & " class=" & Chr(34) & pstrClass & Chr(34) _
						 & " title=" & Chr(34) & Server.HTMLEncode(stripHTML(paryCategories(lngCounter, en_CatFields_Description))) & Chr(34) _
						 & ">" & pstrStartFont & Server.HTMLEncode(paryCategories(lngCounter, en_CatFields_Name)) & pstrEndFont & "</a>"
			Else
				pstrLink = "<a href=" & Chr(34) & pstrTempURL & Chr(34) _
						 & " title=" & Chr(34) & Server.HTMLEncode(stripHTML(paryCategories(lngCounter, en_CatFields_Description))) & Chr(34) _
						 & ">" & pstrStartFont & Server.HTMLEncode(paryCategories(lngCounter, en_CatFields_Name)) & pstrEndFont & "</a>"
			End If
		End With

		CategoryLink = pstrLink
		
	End Function	'CategoryLink
	
	'***********************************************************************************************
	
	Private Function TrailCategoryLink(byVal lngCounter, byVal blnCategory)
	'Input: lngCounter - array pointer
	'	  : blnCategory - True returns category link; False returns subCategory link
	'Output: link
	
	Dim pstrLink
	Dim pstrTempTarget
	
		If CBool(paryCategories(lngCounter, en_CatFields_IsBottom)) Then
			pstrTempTarget = pstrBottomCategoryTargetPage
			If Len(pstrTempTarget) = 0 Then pstrTempTarget = pstrIntermediateCategoryTargetPage
		Else
			pstrTempTarget = pstrIntermediateCategoryTargetPage
			If Len(pstrTempTarget) = 0 Then pstrTempTarget = pstrBottomCategoryTargetPage
		End If

		With pclssfSearch
			.txtsearchParamCat  = paryCategories(lngCounter, en_CatFields_CategoryID)
			.UseShortLink = pblnUseShortLink
			If blnCategory OR paryCategories(lngCounter, en_CatFields_ParentLevel) = 0 Then
				pstrLink = "<a href=" & Chr(34) & pstrTempTarget & "?" & pclssfSearch.SearchLinkParameters(True) & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & Server.HTMLEncode(StripHTML(paryCategories(lngCounter, en_CatFields_Name))) & pstrEndFont & "</a>"
			Else
				.subcat = paryCategories(lngCounter, en_CatFields_uid)
				.iLevel  = cStr(paryCategories(lngCounter, en_CatFields_ParentLevel)+1)
				.txtCatName = cStr(paryCategories(lngCounter, en_CatFields_ParentLevel)+1)
				pstrLink = "<a href=" & Chr(34) & pstrTempTarget & "?" & pclssfSearch.SearchLinkParameters(blnCategory) & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & Server.HTMLEncode(StripHTML(paryCategories(lngCounter, en_CatFields_Name))) & pstrEndFont & "</a>"
			End If
		End With

		TrailCategoryLink = pstrLink
		
	End Function	'TrailCategoryLink	
	
	'***********************************************************************************************
	
	Private Function getArrayPositionByID(ByVal lngCatID)

		Dim i

		For i = 0 To plngNumCategories
			If paryCategories(i, en_CatFields_uid) = lngCatID Then
				getArrayPositionByID = i
				Exit For
			End If
		Next 'i

	End Function 'getArrayPositionByID
	
	'***********************************************************************************************
	
	Private Function getArrayPositionByCategoryID(ByVal lngCatID)

		Dim i

		For i = 0 To plngNumCategories
			If paryCategories(i, en_CatFields_CategoryID) = lngCatID Then
				getArrayPositionByCategoryID = i
				Exit For
			End If
		Next 'i

	End Function 'getArrayPositionByCategoryID
	
	'***********************************************************************************************
	
	Public Function createTrail(ByVal lngCurrentID)

		Dim i
		Dim plngUID
		Dim plngParentID
		Dim pstrURL
		Dim pstrURL_Output
		Dim pstrCategoryName
		Dim plngSafetyCounter
		Dim plngDepth

		On Error Resume Next
		
		If lngCurrentID = -1 Then Exit Function
		
		Call LoadCategories
		
		plngUID = lngCurrentID
		plngSafetyCounter = -1
		plngDepth = -1
		
		If cblnDebugCategorySearchTool Then Response.Write "<h4>creating trail by ID " & plngUID & "</h4>"
		Do While plngDepth <> 0
			i = getArrayPositionByID(plngUID)
			
			If Len(CStr(i)) = 0 Or Not isNumeric(i) Then Exit Do
			If i < 0 Then Exit Do
			
			plngParentID = paryCategories(i, en_CatFields_ParentID)
			pstrCategoryName = paryCategories(i, en_CatFields_Name)
			plngDepth = paryCategories(i, en_CatFields_ParentLevel)
			
			If plngUID <> lngCurrentID Then
				If paryCategories(i, en_CatFields_ParentLevel) = 0 Then
					pstrURL_Output = TrailCategoryLink(i, True) & cstrTrailSpacerText & pstrURL_Output
				Else
					pstrURL_Output = TrailCategoryLink(i, False) & cstrTrailSpacerText & pstrURL_Output
				End If
			ElseIf pblnDisplayCurrentCategoryAsLink Then
				If paryCategories(i, en_CatFields_ParentLevel) = 0 Then
					pstrURL_Output = TrailCategoryLink(i, True) & pstrURL_Output
				Else
					pstrURL_Output = TrailCategoryLink(i, False) & pstrURL_Output
				End If
			ElseIf pblnShowCurrentCategoryAsText Then
				pstrURL_Output = Replace(pstrCurrentCategoryTextWrapper, "{currentCategoryName}", Server.HTMLEncode(StripHTML(paryCategories(i, en_CatFields_Name)))) & pstrURL_Output
			End If
			
			If cblnDebugCategorySearchTool Then
				Response.Write "<ul>i:</b>" & i & "<br />" & vbcrlf
			End If

			plngUID = plngParentID
			plngSafetyCounter = plngSafetyCounter + 1
			If plngSafetyCounter > plngNumCategories Then Exit Do
		Loop
		
		If Len(pstrCategoryTrailShowHome) > 0 Then pstrURL_Output = "<a href=" & Chr(34) & pstrCategoryTrailShowHomeURL & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & pstrCategoryTrailShowHome & pstrEndFont & "</a>" & cstrTrailSpacerText & pstrURL_Output
		createTrail = pstrURL_Output

	End Function	'createTrail

	'***********************************************************************************************
	
	Private Function setTrailStatuses(ByVal lngCurrentID)

	Dim i
	Dim plngUID
	Dim plngPosition
	Dim plngParentID
	Dim plngSafetyCounter
	Dim plngDepth
	Dim pblnFound
	Dim plngCurrentDepth

		If lngCurrentID = -1 Then Exit Function
		
		plngUID = lngCurrentID
		plngSafetyCounter = -1
		plngDepth = -1
		
		If cblnDebugCategorySearchTool Then Response.Write "<h4>creating trail by ID " & plngUID & "</h4>"
		
		'Set initial starting position
		plngPosition = getArrayPositionByID(plngUID)
		plngParentID = paryCategories(plngPosition, en_CatFields_ParentID)
		plngDepth = paryCategories(plngPosition, en_CatFields_ParentLevel)
		
		'Now determine the nodes which are at the same level as the selected category
		'All nodes with the same Parent ID
		If plngParentID <> 0 Then
			plngPosition = getArrayPositionByID(plngParentID) + 1
			For i = plngPosition To plngNumCategories
				If paryCategories(i, en_CatFields_ParentID) = plngParentID Then paryCategories(i, en_CatFields_InTrail) = en_Trail_SameLevel
				plngCurrentDepth = paryCategories(plngPosition, en_CatFields_ParentLevel)
				If plngCurrentDepth = plngDepth Then
					
				ElseIf plngCurrentDepth < plngDepth Then
					'We've stepped up a level so we're done finding categories at this level
					Exit For
				End If
			Next 'i
		End If

		'all nodes one level below current load with this id as a parent id are the first child
		pblnFound = False
		For i = plngPosition + 1 To plngNumCategories
			If paryCategories(i, en_CatFields_ParentID) = plngUID Then
				pblnFound = True
				paryCategories(i, en_CatFields_InTrail) = en_Trail_1stChild
			ElseIf pblnFound Then
				'Only get to this if previously found - this means the section of categories one level down has been passed
				Exit For
			End If
			
		Next 'i
		
		'Now determine nodes above current node which is in the trail
		Do While plngDepth <> 0
			i = getArrayPositionByID(plngUID)

			plngParentID = paryCategories(i, en_CatFields_ParentID)
			plngDepth = paryCategories(i, en_CatFields_ParentLevel)
			
			If plngUID <> lngCurrentID Then
				paryCategories(i, en_CatFields_InTrail) = en_Trail_InTrail
			End If
			
			plngUID = plngParentID
			plngSafetyCounter = plngSafetyCounter + 1
			If plngSafetyCounter > plngNumCategories Then Exit Do
		Loop
		
		If cblnDebugCategorySearchTool Then
			For i = 0 To plngNumCategories
			
				Response.Write "<fieldset style='background-color:white; color=green; font-size:8pt;'><legend>setTrailStatuses (<b>" & i & "</b>)</legend>"
				Response.Write "&nbsp;&nbsp;Name: " & paryCategories(i, en_CatFields_Name) & "<br />"
				Response.Write "&nbsp;&nbsp;uid: " & paryCategories(i, en_CatFields_uid) & "<br />"
				Response.Write "&nbsp;&nbsp;ParentID: " & paryCategories(i, en_CatFields_ParentID) & "<br />"
				Response.Write "&nbsp;&nbsp;ParentLevel: " & paryCategories(i, en_CatFields_ParentLevel) & ""
				Response.Write "&nbsp;&nbsp;InTrail: " & paryCategories(i, en_CatFields_InTrail) & ""
				Response.Write "&nbsp;&nbsp;<hr />"
				Response.Write "&nbsp;&nbsp;plngCurrentUID: " & plngCurrentUID & "<br />"
				Response.Write "&nbsp;&nbsp;plngCurrentDepth: " & plngCurrentDepth & "<br />"
				Response.Write "</fieldset>"
			Next 'i		
		End If
		
	End Function	'setTrailStatuses

	'***********************************************************************************************
	
	Public Function createTrailByCategoryID(ByVal lngCategoryID)

		Dim i
		Dim plngUID

		If lngCategoryID = -1 Then Exit Function
		
		Call LoadCategories
		
		If cblnDebugCategorySearchTool Then Response.Write "<h4>creating trail by CategoryID " & plngUID & "</h4>"
		i = getArrayPositionByCategoryID(lngCategoryID)
		plngUID = paryCategories(i, en_CatFields_uid)

		createTrailByCategoryID = createTrail(plngUID)

	End Function	'createTrailByCategoryID

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

	Public Sub WriteVendorsAsTable

	Dim i
	Dim pclsSFSearch
	Dim plngNumItems
	Dim pstrLineOut
	Dim pstrLink
	Dim pstrTemp
	
		pstrTemp = pstrTemp & "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf

		If loadVendorArray Then
			plngNumItems = UBound(maryVendors)
			For i = 0 To plngNumItems
				pstrLineOut = ""
				If maryVendors(i)(0) <> 1 Then	'Exclude No Vendor
					Set pclsSFSearch = New sfSearch
					pclsSFSearch.txtsearchParamVen = maryVendors(i)(0)
					'Response.Write maryVendors(i)(1) & "<br />"
					pstrLink = "<a href=" & Chr(34) & "search_results.asp" & "?" & pclssfSearch.SearchLinkParameters(False) & Chr(34) & " class=""" & pstrHRefClass & """>" & pstrStartFont & Server.HTMLEncode(maryVendors(i)(1)) & pstrEndFont & "</a>"
					Set pclsSFSearch = Nothing
					pstrLineOut = "<tr><td class=""" & pstrTDClass & """ valign=""middle"">" & pstrLink & "</td></tr>" & vbcrlf
					pstrTemp = pstrTemp & pstrLineOut
				End If
			Next 'i
		End If
		pstrTemp = pstrTemp & "</table>" & vbcrlf
		
		Response.Write pstrTemp
			
	End Sub	'WriteVendorsAsTable

	'***********************************************************************************************
	
End Class	'ssCategoryAE

'***********************************************************************************************
'***********************************************************************************************

	Function CloseListTags(bytNum)

	Dim i
	Dim pstrTemp

		For i = bytNum to 1 Step -1
			pstrTemp = pstrTemp & Space((i-1) * 2) & "</UL>" & vbcrlf
		Next
		
	End Function

	Function Repeat(bytNum,strSource)

	Dim i
	Dim pstrTemp

		For i = 1 to cLng(bytNum)
			pstrTemp = pstrTemp & strSource
		Next
		Repeat = pstrTemp
		
	End Function

	Function WriteHiddenField(strName,strValue)
		WriteHiddenField = "<input type=" & chr(34) & "hidden" & chr(34) & " id=" & chr(34) & strName & chr(34) & " name=" & chr(34) & strName & chr(34) & " value=" & chr(34) & strValue & chr(34) & ">" & vbcrlf
	End Function

	Function WriteOption(strValue,strText,blnSelected)

		If len(strText) = 0 Then strText = strValue
		If blnSelected Then
			WriteOption = "<option value=" & chr(34) & strValue & chr(34) & " selected>" & Server.HTMLEncode(strText) & "</option>" & vbcrlf
		Else
			WriteOption = "<option value=" & chr(34) & strValue & chr(34) & ">" & Server.HTMLEncode(strText) & "</option>" & vbcrlf
		End If
	End Function

	'***********************************************************************************************

'--------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------

Sub WriteSearchFormAE

Dim pclsCategory

	Set pclsCategory = New ssCategoryAE
	pclsCategory.WriteSearchForm(1)
	Set pclsCategory = Nothing
	
End Sub	'WriteSearchFormAE

'***********************************************************************************************

Sub WriteCategoriesAE

Dim pclsCategory

	Set pclsCategory = New ssCategoryAE
	pclsCategory.WriteCategoriesAsTable
	Set pclsCategory = Nothing
	
End Sub	'WriteCategoriesAE

'***********************************************************************************************

Sub WriteSingleSelectAE(byVal strSubCat)

Dim pclsCategory

	Set pclsCategory = New ssCategoryAE
	Response.Write pclsCategory.SelectCategoryAE(strSubCat, 1)
	'Response.Write "<select size=50>" & pclsCategory.getDropdownData & "</select>"

	Response.Write  WriteHiddenField("txtsearchParamCat","ALL")
	'Response.Write  WriteHiddenField("subcat","")
	'Response.Write  WriteHiddenField("txtCatName","")
	Set pclsCategory = Nothing

End Sub	'WriteSingleSelectAE

'***********************************************************************************************

Sub writeCategoryTrailAE(ByVal lngCurrentID, ByVal iLevel, ByVal sSubcat, byRef objclsCategory, byVal blnShowHome)

Dim pclsCategory
Dim pblnLocallyCreated: pblnLocallyCreated = False

	If cblnDebugCategorySearchTool Then
		Response.Write "<table border=""1"" cellspacing=""0"" cellpadding=""2"">"
		Response.Write "<tr><th colspan=2>Sub writeCategoryTrailAE</th></tr>"
		Response.Write "<tr><td>lngCurrentID&nbsp;</td><td>" & lngCurrentID & "&nbsp;</td></tr>"
		Response.Write "<tr><td>iLevel&nbsp;</td><td>" & iLevel & "&nbsp;</td></tr>"
		Response.Write "<tr><td>sSubcat&nbsp;</td><td>" & sSubcat & "&nbsp;</td></tr>"
		Response.Write "<tr><td colspan=2><hr></td></tr>"
		Response.Write "</table>"
	End If

	If isObject(objclsCategory) Then
		Set pclsCategory = objclsCategory
	Else
		pblnLocallyCreated = True
		Set pclsCategory = New ssCategoryAE
	End If
	
	'pclsCategory.BottomCategoryTargetPage = "search_results.asp"
	'pclsCategory.IntermediateCategoryTargetPage = ""
	
	If (sSubcat = lngCurrentID) And iLevel = 2 Then iLevel = 1
	If (sSubcat = lngCurrentID) And iLevel = 0 Then iLevel = 1
	If iLevel = 1 AND (sSubcat = lngCurrentID) Then
		If cblnDebugCategorySearchTool Then
			Response.Write "<tr><th colspan=2>Creating trail with category id</th></tr>"
			Response.Write "<tr><td>Category ID&nbsp;</td><td>" & lngCurrentID & "&nbsp;</td>"
			Response.Write "</table>"
		End If
		If Not blnShowHome Then pclsCategory.CategoryTrailShowHome = ""
		If isNumeric(lngCurrentID) Then Response.Write pclsCategory.createTrailByCategoryID(lngCurrentID)
	Else
		If cblnDebugCategorySearchTool Then
			Response.Write "<tr><th colspan=2>Creating trail with subcategory id</th></tr>"
			Response.Write "<tr><td>subcategory ID&nbsp;</td><td>" & sSubcat & "&nbsp;</td>"
			Response.Write "</table>"
		End If
		If Not blnShowHome Then pclsCategory.CategoryTrailShowHome = ""
		If isNumeric(sSubcat) Then Response.Write pclsCategory.createTrail(sSubcat)
	End If

	If pblnLocallyCreated Then Set pclsCategory = Nothing

End Sub	'writeCategoryTrailAE

'***********************************************************************************************

Sub writeDetailCategoryTrailAE(byVal strProductId, byVal blnShowHome)

Dim pclsCategory
Dim pclsSFSearch
Dim pobjRS
Dim pstrSQL
Dim pstrProductID
Dim plngIDToUse
Dim plngCategoryIDToUse
Dim i
Dim paryIDs

	Set pclsSFSearch = New sfSearch
	pclsSFSearch.LoadSavedSearchCriteria("searchCriteria")
	
	'Protect against SQL Injection
	pstrProductID = Replace(strProductID, "'--", "")
	pstrProductID = Replace(strProductID, "'", "''")
	
	If cblnSF5AE Then
		pstrSQL = "SELECT sfSub_Categories.subcatCategoryId, CatHierarchy" _
				& " FROM (sfProducts INNER JOIN sfsubCatdetail ON sfProducts.prodID = sfsubCatdetail.ProdID) INNER JOIN sfSub_Categories ON sfsubCatdetail.subcatCategoryId = sfSub_Categories.subcatID" _
				& " WHERE (sfProducts.prodID='" & pstrProductID & "')"
		
		pstrSQL = "SELECT sfSub_Categories.subcatID, sfSub_Categories.subcatCategoryId, sfSub_Categories.CatHierarchy" _
				& " FROM sfsubCatdetail LEFT JOIN sfSub_Categories ON sfsubCatdetail.subcatCategoryId = sfSub_Categories.subcatID" _
				& " WHERE (sfsubCatdetail.prodID='" & pstrProductID & "')"
	Else
		pstrSQL = "SELECT prodCategoryId As subcatID, prodCategoryId As subcatCategoryId, prodCategoryId As CatHierarchy" _
				& " FROM sfProducts" _
				& " WHERE (prodID='" & pstrProductID & "')"
	End If 

	If cblnDebugCategorySearchTool Then
		Response.Write "<table border=""1"" cellspacing=""0"" cellpadding=""3"">"
		Response.Write "<tr><th colspan=2>writeDetailCategoryTrail</th></tr>"
		Response.Write "<tr><td>pstrProductID&nbsp;</td><td>" & pstrProductID & "&nbsp;</td>"
		Response.Write "<tr><td>pstrSQL&nbsp;</td><td>" & pstrSQL & "&nbsp;</td>"
	End If

	Set pobjRS = CreateObject("ADODB.Recordset")
	With pobjRS
        .CursorLocation = 2 'adUseClient
		.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
		
		If cblnDebugCategorySearchTool Then
			Response.Write "<tr><td>RecordCount&nbsp;</td><td>" & .RecordCount & "&nbsp;</td>"
		End If
		Do While Not .EOF
			If Len(plngIDToUse) = 0 Then
				plngIDToUse = Trim(.Fields("CatHierarchy").Value & "")
				plngCategoryIDToUse = Trim(.Fields("subcatCategoryId").Value & "")
			End If
			
			If cblnDebugCategorySearchTool Then
				Response.Write "<tr><td><b>plngIDToUse</b>&nbsp;</td><td>" & plngIDToUse & "&nbsp;</td>"
				Response.Write "<tr><td>pclsSFSearch.subcat&nbsp;</td><td>" & pclsSFSearch.subcat & "&nbsp;</td>"
				Response.Write "<tr><td>CatHierarchy&nbsp;</td><td>" & .Fields("CatHierarchy").Value & "&nbsp;</td>"
			End If

			'Now if there isn't a search in storage then use the first category match
			If pclsSFSearch.HasStoredSearch Then
				paryIDs = Split(.Fields("CatHierarchy").Value, "-")
				For i = 0 To UBound(paryIDs)
					'Response.Write "<tr><td>" & i & "&nbsp;</td><td>" & paryIDs(i) & "&nbsp;</td>"
					If paryIDs(i) = pclsSFSearch.subcat Then
						plngIDToUse = paryIDs(i)
						plngCategoryIDToUse = Trim(.Fields("subcatCategoryId").Value & "")
						Exit Do
					End If
				Next 'i
			Else
				Exit Do
			End If
			

			.MoveNext
		Loop
		
		If InStr(1, plngIDToUse, "-") > 0 Then
			paryIDs = Split(plngIDToUse, "-")
			plngIDToUse = paryIDs(UBound(paryIDs))
		End If
		
		If cblnDebugCategorySearchTool Then
			Response.Write "<tr><td colspan=2><hr></td>"
			Response.Write "<tr><td>plngIDToUse&nbsp;</td><td>" & plngIDToUse & "&nbsp;</td>"
			Response.Write "<tr><td>plngCategoryIDToUse&nbsp;</td><td>" & plngCategoryIDToUse & "&nbsp;</td>"
			Response.Write "<tr><td>Level&nbsp;</td><td>" & i+2 & "&nbsp;</td>"
			Response.Write "</table>"
		End If
		.Close
	End With
	Set pobjRS = Nothing

	Set pclsCategory = New ssCategoryAE
	pclsCategory.LoadSearchSettings
	If Not blnShowHome Then pclsCategory.CategoryTrailShowHome = ""
	pclsCategory.ShowCurrentCategoryAsText = False
	pclsCategory.DisplayCurrentCategoryAsLink = True
	Call writeCategoryTrail(plngCategoryIDToUse, i+2, plngIDToUse, pclsCategory, blnShowHome)
	Set pclsCategory = Nothing
	Set pclsSFSearch = Nothing

End Sub	'writeDetailCategoryTrailAE

%>