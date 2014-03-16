<%
'/////////////////////////////////////////////////
'/
'/  User Parameters
'/


'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'Possible Content Types
Const enContentType_Breeds = 9
Const enContentType_Designs = 11
Const enContentType_Themes = 10

'**********************************************************
'*	Page Level variables
'**********************************************************

Const CMS_CacheDuration = 0

Dim maryCMS
Dim mlngContentID
Dim mstrContentType

'**********************************************************
'*	Functions
'**********************************************************

'Sub DeterminePageIntent()
'Sub displayCMSItemContent()
'Function getContentElement(byVal strElement)
'Function loadContentByID(byVal lngContentID)
'Function loadContentByURL(byVal strURL)
'Function loadContentCategories(byRef aryContent)
'Sub setContentToArray(byRef objRS)

'**********************************************************
'*	Begin Page Code
'**********************************************************

'**********************************************************
'*******************************************************************************************************

Sub DeterminePageIntent(byRef lngPageNumber)

Dim paryTemp
Dim plngPos
Dim plngPos2
Dim pstrTemp
Dim pstrTempPageNumber
Dim pstrURL

'possible querystrings
'nothing
'contentType
'
	pstrTemp = Request.QueryString
	If Len(pstrTemp) > 4 Then
		If Left(pstrTemp, 4) = "404;" Then
			pstrURL = Replace(LCase(pstrTemp), "404;" & LCase(adminDomainName), "")
			mstrContentType = "UNKNOWN"	'Set to default value - will be changed later if match found
			
			plngPos = InStrRev(pstrTemp, "&")
			If plngPos > 1 Then
				paryTemp = Split(Right(pstrTemp, Len(pstrTemp) - plngPos), "=")
				
				If UBound(paryTemp) >= 1 Then
					mlngContentID = Trim(paryTemp(1))

					'Now determine the page number if any
					plngPos = InStrRev(pstrTemp, "/")
					If plngPos > 1 Then
						plngPos2 = InStrRev(pstrTemp, "/", plngPos - 1)
						If plngPos2 > 1 Then
							pstrTempPageNumber = Mid(pstrTemp, plngPos2, plngPos - plngPos2)
							If Len(pstrTempPageNumber) > 5 Then
								pstrTempPageNumber = Replace(LCase(pstrTempPageNumber), "/page", "")
								If isNumeric(pstrTempPageNumber) Then lngPageNumber = CLng(pstrTempPageNumber)
							End If
						End If
					End If

					If isNumeric(mlngContentID) And Len(mlngContentID) > 0 Then
						Select Case paryTemp(0)
							Case "cmsID", "contentType", "mfgID"
								mstrContentType = paryTemp(0)
							Case Else
								If loadContentByURL(pstrURL) Then
								Else
									Response.Write "Unknown content type: " & paryTemp(0) & "<br />"
									mstrContentType = "UNKNOWN"
								End If
						End Select
					Else
						mlngContentID = ""
					End If	'isNumeric(mlngContentID) And Len(mlngContentID) > 0
				End If	'UBound(paryTemp) >= 1
			Else
				If loadContentByURL(pstrURL) Then

				End If
			End If	'plngPos > 1
		Else
			'For Testing
			If Len(LoadRequestValue("cmsID")) > 0 Then
				mstrContentType = "cmsID"
				mlngContentID = LoadRequestValue("cmsID")
			ElseIf Len(LoadRequestValue("mfgID")) > 0 Then
				mstrContentType = "mfgID"
				mlngContentID = LoadRequestValue("mfgID")
			ElseIf Len(LoadRequestValue("contentType")) > 0 Then
				mstrContentType = "contentType"
				mlngContentID = LoadRequestValue("contentType")
			Else
				mstrContentType = "UNKNOWN"
			End If
		End If	'Left(pstrTemp, 4) = "404;"
	End If	'Len(pstrTemp) > 4
	
	If cblnDebugCMS Then
		Response.Write "<fieldset><legend>DeterminePageIntent (ssl/SFLib/ssmodCMS.asp)</legend>"
		Response.Write "ContentType: " & mstrContentType & "<br />"
		Response.Write "ContentID: " & mlngContentID & "<br />"
		Response.Write "PageNumber: " & lngPageNumber & "<hr />"
		Response.Write "Querystring: " & pstrTemp & "<br />"
		Response.Write "HTTP_REFERER: " & Request.ServerVariables("HTTP_REFERER") & "<br />"
		Response.Write "URL: " & pstrURL
		Response.Write "</fieldset>"
	End If

End Sub	'DeterminePageIntent

'**********************************************************

Sub displayCMSItemContent()

	Response.Write "<fieldset><legend>Content for <em>" & getContentElement("contentID") & "</em></legend>"
	Response.Write "contentAuthorID: " & getContentElement("AuthorID") & "<br />"
	Response.Write "contentContentType: " & getContentElement("ContentType") & "<br />"
	Response.Write "contentReferenceID: " & getContentElement("ReferenceID") & "<br />"
	Response.Write "contentApprovedForDisplay: " & getContentElement("ApprovedForDisplay") & "<br />"
	Response.Write "contentAbstract: " & getContentElement("Abstract") & "<br />"
	Response.Write "contentContent: " & getContentElement("Content") & "<br />"
	Response.Write "contentContentFilePath: " & getContentElement("ContentFilePath") & "<br />"
	Response.Write "contentAuthorName: " & getContentElement("AuthorName") & "<br />"
	Response.Write "contentAuthorEmail: " & getContentElement("AuthorEmail") & "<br />"
	Response.Write "contentAuthorShowEmail: " & getContentElement("AuthorShowEmail") & "<br />"
	Response.Write "contentAuthorRating: " & getContentElement("AuthorRating") & "<br />"
	Response.Write "contentDateCreated: " & getContentElement("DateCreated") & "<br />"
	Response.Write "contentDateModified: " & getContentElement("DateModified") & "<br />"
	Response.Write "contentTemplatePage: " & getContentElement("TemplatePage") & "<br />"
	Response.Write "contentPageName: " & getContentElement("PageName") & "<br />"
	Response.Write "contentPageTitle: " & getContentElement("PageTitle") & "<br />"
	Response.Write "contentMetaDescription: " & getContentElement("MetaDescription") & "<br />"
	Response.Write "contentMetaKeywords: " & getContentElement("MetaKeywords") & "<br />"
	Response.Write "contentMetaAuthor: " & getContentElement("MetaAuthor") & "<br />"
	Response.Write "contentMetaCustom2: " & getContentElement("MetaCustom2") & "<br />"
	Response.Write "contentSortOrder: " & getContentElement("SortOrder") & "<br />"
	Response.Write "contentTitle: " & getContentElement("Title") & "<br />"
	Response.Write "</fieldset>"
	
End Sub	'displayCMSItemContent

'**********************************************************

Function getContentElement(byVal strElement)

	If isArray(maryCMS) Then
		Select Case strElement
			Case "AuthorID":			getContentElement = maryCMS(0)
			Case "ContentType":			getContentElement = maryCMS(1)
			Case "ReferenceID":			getContentElement = maryCMS(2)
			Case "ApprovedForDisplay":	getContentElement = maryCMS(3)
			Case "Title":				getContentElement = maryCMS(4)
			Case "Abstract":			getContentElement = maryCMS(5)
			Case "Content":				getContentElement = maryCMS(6)
			Case "ContentFilePath":		getContentElement = maryCMS(7)
			Case "AuthorName":			getContentElement = maryCMS(8)
			Case "AuthorEmail":			getContentElement = maryCMS(9)
			Case "AuthorShowEmail":		getContentElement = maryCMS(10)
			Case "AuthorRating":		getContentElement = maryCMS(11)
			Case "DateCreated":			getContentElement = maryCMS(12)
			Case "DateModified":		getContentElement = maryCMS(13)
			Case "TemplatePage":		getContentElement = maryCMS(14)
			Case "PageName":			getContentElement = maryCMS(15)
			Case "PageTitle":			getContentElement = maryCMS(16)
			Case "MetaDescription":		getContentElement = maryCMS(17)
			Case "MetaKeywords":		getContentElement = maryCMS(18)
			Case "MetaAuthor":			getContentElement = maryCMS(19)
			Case "MetaCustom1":			getContentElement = maryCMS(20)
			Case "MetaCustom2":			getContentElement = maryCMS(21)
			Case "SortOrder":			getContentElement = maryCMS(22)
			Case "contentID":			getContentElement = mlngContentID
		End Select
	End If	'isArray(maryCMS)
	
End Function	'getContentElement

'**********************************************************

Sub resetContent(byVal lngContentID)
	If CStr(mlngContentID) <> CStr(lngContentID) Then Set maryCMS = Nothing
End Sub

'**********************************************************

Function loadContentByID(byVal lngContentID)

Dim pblnSuccess
Dim pobjCmd
Dim pobjRS

	pblnSuccess = False
	If Len(lngContentID) = 0 Or Not isNumeric(lngContentID) Then
		loadContentByID = pblnSuccess
		Exit Function
	End If
	
	Call resetContent(lngContentID)
	If isArray(maryCMS) Then
		loadContentByID = True
		Exit Function
	End If
	
	'Call DebugRecordSplitTime("Loading content . . .")
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = "Select contentID, contentAuthorID, contentContentType, contentReferenceID, contentApprovedForDisplay, contentTitle, contentAbstract, contentContent, contentContentFilePath, contentAuthorName, contentAuthorEmail, contentAuthorShowEmail, contentAuthorRating, contentDateCreated, contentDateModified, contentTemplatePage, contentPageName, contentPageTitle, contentMetaDescription, contentMetaKeywords, contentMetaAuthor, contentMetaCustom1, contentMetaCustom2, contentSortOrder" _
						& " From content" _
						& " Where contentID=?"
		Set .ActiveConnection = cnn
		
		'On Error Resume Next
		.Parameters.Append .CreateParameter("contentID", adInteger, adParamInput, 4, lngContentID)
		Set pobjRS = .Execute
		If Err.number <> 0 Then
			loadContentByID = False
			Err.Clear
			Exit Function
		End If

		If Not pobjRS.EOF Then
			Call setContentToArray(pobjRS)
			pblnSuccess = True
		End If
		
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCmd
	Set pobjCmd = Nothing
	Call DebugRecordSplitTime("Content loaded")

	loadContentByID = pblnSuccess
	
End Function	'loadContentByID

'**********************************************************

Function loadContentByURL(byVal strURL)

Dim pblnFound
Dim pobjCmd
Dim pobjRS
Dim paryContent
Dim pstrCacheKey

  	pblnFound = False
	If Len(strURL) > 0 Then
		pstrCacheKey = "cmsPage_" & strURL
		Application.Contents.Remove(pstrCacheKey)
		paryContent = getFromCache(pstrCacheKey)
		If isArray(paryContent) Then
			maryCMS = paryContent
		Else
			Set pobjCmd  = CreateObject("ADODB.Command")
			With pobjCmd
				.Commandtype = adCmdText
				.Commandtext = "SELECT * FROM content WHERE contentPageName=?"
				Set .ActiveConnection = cnn

				.Parameters.Append .CreateParameter("key", adVarChar, adParamInput, Len(strURL), strURL)
  				Set pobjRS = .Execute
  				If Not pobjRS.EOF Then
  					If ConvertToBoolean(pobjRS.Fields("contentApprovedForDisplay").Value, False) Then
  						pblnFound = True
  						Call setContentToArray(pobjRS)
						Call saveToCache(pstrCacheKey, maryCMS, DateAdd("s", 600, Now()))
  					End If
  				End If
  				closeObj(pobjRS)
			End With
		End If	'Len(pstrContent) = 0
		Set pobjCmd = Nothing
	
	End If	'Len(strURL) > 0

	loadContentByURL = pblnFound

End Function	'loadContentByURL

'**********************************************************

Function loadContentCategories(byRef aryContent)

Dim pblnSuccess
Dim pobjRS
Dim pstrSQL

	pblnSuccess = False
	
	Call DebugRecordSplitTime("Loading content categories . . .")
	pstrSQL = "Select contentTypeID, contentTypeDisplayName, contentTypeURL, contentTypeDescription From contentTypes Where contentTypeDisplayInSiteMap<>0 Order By contentTypeDisplayOrder, contentTypeDisplayName"
	Set pobjRS=GetRS(pstrSQL)
	With pobjRS
		If Not .EOF Then
			aryContent = .GetRows
			pblnSuccess = True
		End If
	End With
	Call closeObj(pobjRS)
	Call DebugRecordSplitTime("Content categories loaded")
	
	loadContentCategories = pblnSuccess
	
End Function	'loadContentCategories

'**********************************************************

Function loadContentTypeByID(byVal lngContentType, byRef aryContent)

Dim pblnSuccess
Dim pobjRS
Dim pstrSQL

	pblnSuccess = False
	
	Call DebugRecordSplitTime("Loading content categories . . .")
	pstrSQL = "Select contentTypeID, contentTypeDisplayName, contentTypeURL, contentTypeDescription From contentTypes Where contentTypeID=" & lngContentType
	Set pobjRS=GetRS(pstrSQL)
	With pobjRS
		If Not .EOF Then
			aryContent = .GetRows
			pblnSuccess = True
		End If
	End With
	Call closeObj(pobjRS)
	Call DebugRecordSplitTime("Content categories loaded")
	
	loadContentTypeByID = pblnSuccess
	
End Function	'loadContentTypeByID

'**********************************************************

Sub setContentToArray(byRef objRS)

	With objRS
		If Not .EOF Then
			ReDim maryCMS(23)
			maryCMS(0) = Trim(.Fields("contentAuthorID").Value & "")
			maryCMS(1) = Trim(.Fields("contentContentType").Value & "")
			maryCMS(2) = Trim(.Fields("contentReferenceID").Value & "")
			maryCMS(3) = Trim(.Fields("contentApprovedForDisplay").Value & "")
			maryCMS(4) = Trim(.Fields("contentTitle").Value & "")
			maryCMS(5) = Trim(.Fields("contentAbstract").Value & "")
			maryCMS(6) = Trim(.Fields("contentContent").Value & "")
			maryCMS(7) = Trim(.Fields("contentContentFilePath").Value & "")
			maryCMS(8) = Trim(.Fields("contentAuthorName").Value & "")
			maryCMS(9) = Trim(.Fields("contentAuthorEmail").Value & "")
			maryCMS(10) = Trim(.Fields("contentAuthorShowEmail").Value & "")
			maryCMS(11) = Trim(.Fields("contentAuthorRating").Value & "")
			maryCMS(12) = Trim(.Fields("contentDateCreated").Value & "")
			maryCMS(13) = Trim(.Fields("contentDateModified").Value & "")
			maryCMS(14) = Trim(.Fields("contentTemplatePage").Value & "")
			maryCMS(15) = Trim(.Fields("contentPageName").Value & "")
			maryCMS(16) = Trim(.Fields("contentPageTitle").Value & "")
			maryCMS(17) = Trim(.Fields("contentMetaDescription").Value & "")
			maryCMS(18) = Trim(.Fields("contentMetaKeywords").Value & "")
			maryCMS(19) = Trim(.Fields("contentMetaAuthor").Value & "")
			maryCMS(20) = Trim(.Fields("contentMetaCustom1").Value & "")
			maryCMS(21) = Trim(.Fields("contentMetaCustom2").Value & "")
			maryCMS(22) = Trim(.Fields("contentSortOrder").Value & "")
			mlngContentID = .Fields("contentID").Value
			'Call displayCMSItemContent
		End If	'.EOF
	End With
	
End Sub	'setContentToArray

'**********************************************************
' Psuedo Content Properties
'**********************************************************

Function contentContentType()
	contentContentType = getContentElement("ContentType")
End Function

Function contentMetaDescription()
	contentMetaDescription = getContentElement("MetaDescription")
End Function

Function contentReferenceID()
	contentReferenceID = getContentElement("ReferenceID")
End Function

Function contentApprovedForDisplay()
	contentApprovedForDisplay = getContentElement("ApprovedForDisplay")
End Function

Function contentAbstract()
	contentAbstract = getContentElement("Abstract")
End Function

Function contentContent()
	contentContent = getContentElement("Content")
End Function

Function contentContentFilePath()
	contentContentFilePath = getContentElement("ContentFilePath")
End Function

Function contentAuthorName()
	contentAuthorName = getContentElement("AuthorName")
End Function

Function contentAuthorEmail()
	contentAuthorEmail = getContentElement("AuthorEmail")
End Function

Function contentAuthorShowEmail()
	contentAuthorShowEmail = getContentElement("AuthorShowEmail")
End Function

Function contentAuthorRating()
	contentAuthorRating = getContentElement("AuthorRating")
End Function

Function contentDateCreated()
	contentDateCreated = getContentElement("DateCreated")
End Function

Function contentDateModified()
	contentDateModified = getContentElement("DateModified")
End Function

Function contentTemplatePage()
	contentTemplatePage = getContentElement("TemplatePage")
End Function

Function contentPageName()
	contentPageName = getContentElement("PageName")
End Function

Function contentPageTitle()
	contentPageTitle = getContentElement("PageTitle")
End Function

Function contentTitle()
	contentTitle = getContentElement("Title")
End Function

Function contentMetaDescription()
	contentMetaDescription = getContentElement("MetaDescription")
End Function

Function contentMetaKeywords()
	contentMetaKeywords = getContentElement("MetaKeywords")
End Function

Function contentMetaAuthor()
	contentMetaAuthor = getContentElement("MetaAuthor")
End Function

Function contentMetaCustom1()
	contentMetaCustom1 = getContentElement("MetaCustom1")
End Function

Function contentMetaCustom2()
	contentMetaCustom2 = getContentElement("MetaCustom2")
End Function

Function contentSortOrder()
	contentSortOrder = getContentElement("SortOrder")
End Function

'**********************************************************
'**********************************************************

Class clsXMLCMS

Dim p_objxmlDoc
Dim p_objxmlRoot
Dim p_objxmlNode

Dim p_strXMLOut

Dim p_aryFieldDetails(6)

Dim plngID
Dim plngPageCount
Dim plngPageSize
Dim plngPageNumber
Dim plngRecordCount
Dim pstrTemplatePath
Dim querystring
Dim pstrtext
Dim pstrTitle
Dim pstrhref
Dim pstrCount
'0) ID		- ID - from appropriate table, used for querystring name/value pair													-
'1) querystring		- namefrom appropriate table, used for querystring name/value pair
'2) text		- Display Text														- 
'3) title		- How do you want the display										- enDisplayType_hidden, enDisplayType_select, enDisplayType_textarea, enDisplayType_textbox, enDisplayType_checkbox, enDisplayType_listbox, enDisplayType_textbox_WithDateSelect, enDisplayType_textbox_WithHTMLSelect
'4) href	-
'5) count - # items in this category
'6) image

'***********************************************************************************************

Private Sub class_Initialize()
	plngPageSize = 10
	plngPageNumber = 1
	plngPageCount = 0
	plngRecordCount = 0
End Sub

Private Sub class_Terminate()

    On Error Resume Next
    
	Call ReleaseObject(p_objxmlDoc)
	Call ReleaseObject(p_objxmlRoot)
	Call ReleaseObject(p_objxmlNode)

End Sub

'***********************************************************************************************

Public Property Let name(vData)
	p_aryFieldDetails(0) = vData
End Property

Public Property Let value(vData)
	p_aryFieldDetails(1) = vData
End Property

Public Property Let text(vData)
	p_aryFieldDetails(2) = vData
End Property

Public Property Let title(vData)
	p_aryFieldDetails(3) = vData
End Property

Public Property Let href(vData)
	p_aryFieldDetails(4) = vData
End Property

Public Property Let count(vData)
	p_aryFieldDetails(5) = vData
End Property

Public Property Let image(vData)
	p_aryFieldDetails(6) = vData
End Property

Public Property Let TemplatePath(vData)
	pstrTemplatePath = vData
End Property

'***********************************************************************************************

Public Property Get PageCount
	PageCount = plngPageCount
End Property

Public Property Let PageSize(vData)
	plngPageSize = vData
End Property

Public Property Let PageNumber(vData)
	plngPageNumber = vData
End Property

Public Property Get RecordCount
	RecordCount = plngRecordCount
End Property

'***********************************************************************************************
Private Sub Initialize

	set p_objxmlDoc = CreateObject("MSXML2.DOMDocument")
	
	' Create processing instruction and document root
    Set p_objxmlNode = p_objxmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Set p_objxmlNode = p_objxmlDoc.insertBefore(p_objxmlNode, p_objxmlDoc.childNodes.Item(0))
    
	' Create document root
    Set p_objxmlRoot = p_objxmlDoc.createElement("fields")
    Set p_objxmlDoc.documentElement = p_objxmlRoot
    p_objxmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
    
End Sub	'Initialize

'***********************************************************************************************

Public Sub addItem(byVal strName, byVal strValue, byVal strtext, byVal strTitle, byVal strhref, byVal strImage, byVal strCount)

Dim pstrtext

	pstrtext = Trim(strtext & "")
	If Len(pstrtext) = 0 Then pstrtext = Trim(strName & "")
	If Len(pstrtext) = 0 Then pstrtext = Trim(strTitle & "")
	
	p_aryFieldDetails(0) = strName
	p_aryFieldDetails(1) = strValue
	p_aryFieldDetails(2) = pstrtext
	
	If Len(strTitle) = 0 Then
		p_aryFieldDetails(3) = Server.HTMLEncode(Replace(pstrtext, Chr(34), ""))
	Else
		p_aryFieldDetails(3) = Server.HTMLEncode(Replace(Trim(strTitle & ""), Chr(34), ""))
	End If
	
	If Len(strhref & "") = 0 Then
		p_aryFieldDetails(4) = Server.URLEncode(Replace(pstrtext, Chr(34), "")) & ".htm"
	Else
		p_aryFieldDetails(4) = strhref
	End If
	
	p_aryFieldDetails(5) = strCount
	p_aryFieldDetails(6) = strImage
	
	Call setFieldValues
	
End Sub	'addItem

'***********************************************************************************************

Public Sub setFieldValues

Dim p_objxmlElement
Dim p_objxmlFieldDetail
Dim i

	If Not isObject(p_objxmlDoc) Then Call Initialize

	'Create the field node
    Set p_objxmlFieldDetail = p_objxmlDoc.createElement("field")
    
    'Attach the field not to the fields nodes
    p_objxmlRoot.appendChild p_objxmlFieldDetail
    
	'start hanging on the field elements
	For i = 0 To UBound(p_aryFieldDetails)
		Select Case i
			Case 0	'name
				Set p_objxmlElement = p_objxmlDoc.createElement("name")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
			Case 1	'value
				Set p_objxmlElement = p_objxmlDoc.createElement("value")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
			Case 2	'text
				Set p_objxmlElement = p_objxmlDoc.createElement("text")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
			Case 3	'title
				Set p_objxmlElement = p_objxmlDoc.createElement("title")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
			Case 4	'href
				Set p_objxmlElement = p_objxmlDoc.createElement("href")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
			Case 5	'count
				Set p_objxmlElement = p_objxmlDoc.createElement("count")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
			Case 6	'image
				Set p_objxmlElement = p_objxmlDoc.createElement("image")
				p_objxmlElement.Text = Trim(p_aryFieldDetails(i) & "")
		End Select

		p_objxmlFieldDetail.appendChild p_objxmlElement

	Next 'i

	Set p_objxmlElement = Nothing
	Set p_objxmlFieldDetail = Nothing

End Sub	'setFieldValues

'***********************************************************************************************

Public Sub SetOutputArray(byRef aryOutput)
'This maps the XML document to the array expected for the output

Dim i, j
Dim paryMapping 
Dim paryEmpty

paryMapping = Array("name",	"value",	"text",	"title",	"href",	"count")
paryEmpty	= Array("",		"",			"",		"",			"",		"")

	ReDim aryOutput(p_objxmlRoot.childNodes.length - 1)
	
	For i = 0 To UBound(aryOutput)
		aryOutput(i) = paryEmpty
		For j = 0 To UBound(aryOutput(i))
			If Len(paryMapping(j)) > 0 Then aryOutput(i)(j) = p_objxmlRoot.childNodes(i).selectSingleNode(paryMapping(j)).text
		Next 'j
	Next 'i
	
End Sub	'SetOutputArray

'***********************************************************************************************

Public Property Get XMLText

	If isObject(p_objxmlDoc) Then
		XMLText = p_objxmlDoc.xml
	Else
		XMLText = "No XML Document"
	End If

End Property

'***********************************************************************************************

Public Sub saveXML(byVal strPath)
	If isObject(p_objxmlDoc) Then p_objxmlDoc.save strPath
End Sub

'***********************************************************************************************

Public Function loadXML(byVal strPath)
	Call Initialize
	loadXML = p_objxmlDoc.Load(strPath)
End Function

'*******************************************************************************************************

Public Function isValidCMSFile(byVal strFilePath, byVal cacheDuration)

Dim fso, f
Dim pblnResult

	pblnResult = False
	Set fso = CreateObject("Scripting.FileSystemObject")

	On Error Resume Next

	Set f = fso.GetFile(strFilePath)
	If Err.number = 0 Then
		'debugprint "f.DateLastModified", f.DateLastModified
		pblnResult = CBool(dateAdd("m", cacheDuration, f.DateLastModified) > Now())
	Else
		Err.Clear
	End If

   'DateLastModified
   
   isValidCMSFile = pblnResult

End Function	'isValidCMSFile

'******************************************************************************************************************************************************************

Function getProductsByCMSID(byVal lngCMSID)

Dim i
Dim paryResult
Dim pblnValidData
Dim pbytKeyIndex
Dim plngCacheDuration
Dim pobjRS
Dim pstrCMSFileLocation
Dim pstrGeneralSQL

	If Not isNumeric(lngCMSID) Or Len(lngCMSID) = 0 Then
		getProductsByCMSID = False
		Exit Function
	End If

	pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\ProductsByCMSID_" & lngCMSID & ".xml"
	plngCacheDuration = CMS_CacheDuration
	pbytKeyIndex = 0

	If isValidCMSFile(pstrCMSFileLocation, plngCacheDuration) Then
		pblnValidData = loadXML(pstrCMSFileLocation)
	Else
		pblnValidData = False
	End If

	If Not pblnValidData Then
	
		pstrGeneralSQL = "SELECT sfProducts.prodID, sfProducts.prodLink, sfProducts.metaTitle, sfProducts.prodName, sfProducts.metaDescription, sfProducts.prodShortDescription, sfProducts.prodImageSmallPath" _
					   & " FROM sfProducts INNER JOIN contentProductAssignments ON sfProducts.sfProductID = contentProductAssignments.contentProductAssignmentProductID" _
					   & " WHERE sfProducts.prodEnabledIsActive=1 AND contentProductAssignments.contentProductAssignmentContentID=" & lngCMSID
		
		set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 3			'adUseClient
			.Open pstrGeneralSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
			'Set .Connection = Nothing	'Disconnect the recordset

			If Not .EOF Then
				paryResult = .GetRows()
				.Close
				
				For i = 0 To UBound(paryResult, 2)

					If False Then
						Response.Write "<fieldset><legend>getCMSData</legend>"
						Response.Write "0: " & paryResult(0,i) & "<BR>"
						Response.Write "1: " & paryResult(1,i) & "<BR>"
						Response.Write "2: " & paryResult(2,i) & "<BR>"
						Response.Write "3: " & paryResult(3,i) & "<BR>"
						Response.Write "4: " & paryResult(4,i) & "<BR>"
						Response.Write "5: " & paryResult(5,i) & "<BR>"
						Response.Write "6: " & paryResult(6,i) & "<BR>"
						Response.Write "</fieldset>"
					End If
					
					'check name
					If Len(Trim(paryResult(1,i) & "")) = 0 Then paryResult(1,i) = "detail.asp?product_ID=" & paryResult(0,i)	'use link if present, otherwise use detail.asp
					If Len(Trim(paryResult(2,i) & "")) = 0 Then paryResult(2,i) = paryResult(3,i)	'use metaTitle if present, otherwise use product name
					If Len(Trim(paryResult(4,i) & "")) = 0 Then paryResult(4,i) = paryResult(5,i)	'use metaDescription if present, otherwise use prodShortDescription
					If Len(Trim(paryResult(4,i) & "")) = 0 Then paryResult(4,i) = paryResult(2,i)	'use metaDescription if present, otherwise use name
					
					addItem paryResult(3,i), paryResult(0,i), paryResult(2,i), paryResult(4,i), paryResult(1,i), paryResult(6,i), ""
					'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref, strImage, strCount

				Next
			Else
				.Close
			End If
				
		End With	'pobjRS
		Set	pobjRS = Nothing
		
		If plngCacheDuration > 0 Then saveXML pstrCMSFileLocation

	End If	'pblnValidData
	
	getProductsByCMSID = True
	'Response.Write Server.HTMLEncode(XMLText)

End Function	'getProductsByCMSID

'******************************************************************************************************************************************************************

Function getCMSDataByCategoryID(byVal lngCategoryID, byVal lngContentType)

Dim i
Dim paryResult
Dim pblnValidData
Dim pbytKeyIndex
Dim plngCacheDuration
Dim pobjRS
Dim pstrCMSFileLocation
Dim pstrGeneralSQL

	If Not isNumeric(lngCategoryID) Or Len(lngCategoryID) = 0 Or Not isNumeric(lngContentType) Or Len(lngContentType) = 0 Then
		getCMSDataByReferenceID = False
		Exit Function
	End If

	pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\CMSDataByCategoryID_" & lngContentType & "_" & lngCategoryID & ".xml"
	plngCacheDuration = CMS_CacheDuration
	pbytKeyIndex = 0

	If isValidCMSFile(pstrCMSFileLocation, plngCacheDuration) Then
		pblnValidData = loadXML(pstrCMSFileLocation)
	Else
		pblnValidData = False
	End If

	If Not pblnValidData Then
	
		pstrGeneralSQL = "SELECT contentPageName, contentID, contentTitle, contentContent, contentTemplatePage, contentContentFilePath" _
					   & " FROM content INNER JOIN contentCategoryAssignments ON [content].contentID = contentCategoryAssignments.contentCategoryAssignmentContentID" _
					   & " WHERE ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1)) AND contentContentType=" & lngContentType & " AND contentCategoryAssignmentCategoryID =" & lngCategoryID _
					   & " Order By contentSortOrder, contentTitle Asc"

		set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 3			'adUseClient
			.Open pstrGeneralSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
			'Set .Connection = Nothing	'Disconnect the recordset
Response.Write "pstrGeneralSQL: " & pstrGeneralSQL & "<BR>"
Response.Write ".EOF: " & .EOF & "<BR>"
			If Not .EOF Then
				paryResult = .GetRows()
				.Close
				
				For i = 0 To UBound(paryResult, 2)
					addItem paryResult(0,i), paryResult(1,i), paryResult(2,i), paryResult(3,i), paryResult(4,i), paryResult(5,i), ""
					'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref, strImage, strCount
				Next
			Else
				.Close
			End If
				
		End With	'pobjRS
		Set	pobjRS = Nothing
		
		If plngCacheDuration > 0 Then saveXML pstrCMSFileLocation

	End If	'pblnValidData
	
	getCMSDataByCategoryID = True
	'Response.Write Server.HTMLEncode(XMLText)

End Function	'getCMSDataByCategoryID

'******************************************************************************************************************************************************************

Function getCMSDataByReferenceID(byVal lngReferenceID, byVal lngContentType)

Dim i
Dim paryResult
Dim pblnValidData
Dim pbytKeyIndex
Dim plngCacheDuration
Dim pobjRS
Dim pstrCMSFileLocation
Dim pstrGeneralSQL

	If Not isNumeric(lngReferenceID) Or Len(lngReferenceID) = 0 Or Not isNumeric(lngContentType) Or Len(lngContentType) = 0 Then
		getCMSDataByReferenceID = False
		Exit Function
	End If

	pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\CMSDataByReferenceID_" & lngContentType & "_" & lngReferenceID & ".xml"
	plngCacheDuration = CMS_CacheDuration
	pbytKeyIndex = 0

	If isValidCMSFile(pstrCMSFileLocation, plngCacheDuration) Then
		pblnValidData = loadXML(pstrCMSFileLocation)
	Else
		pblnValidData = False
	End If

	If Not pblnValidData Then
	
		pstrGeneralSQL = "SELECT contentPageName, contentID, contentTitle, contentContent, contentTemplatePage, contentContentFilePath" _
					   & " FROM content" _
					   & " WHERE ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1)) AND contentContentType=" & lngContentType & " AND contentReferenceID=" & lngReferenceID _
					   & " Order By contentSortOrder, contentTitle Asc"

		set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 3			'adUseClient
			.Open pstrGeneralSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
			'Set .Connection = Nothing	'Disconnect the recordset

			If Not .EOF Then
				paryResult = .GetRows()
				.Close
				
				For i = 0 To UBound(paryResult, 2)
					addItem paryResult(0,i), paryResult(1,i), paryResult(2,i), paryResult(3,i), paryResult(4,i), paryResult(5,i), ""
					'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref, strImage, strCount
				Next
			Else
				.Close
			End If
				
		End With	'pobjRS
		Set	pobjRS = Nothing
		
		If plngCacheDuration > 0 Then saveXML pstrCMSFileLocation

	End If	'pblnValidData
	
	getCMSDataByReferenceID = True
	'Response.Write Server.HTMLEncode(XMLText)

End Function	'getCMSDataByReferenceID

'******************************************************************************************************************************************************************

Function getCMSData(byVal lngContentType, byVal strCMSFileLocation)

Dim i
Dim paryResult
Dim pblnValidData
Dim pdicItems
Dim pobjRS
Dim pstrContentSQL
Dim pstrGeneralSQL
Dim pstrKey
Dim pbytKeyIndex

Dim plngCacheDuration

	Select Case lngContentType
		Case 3	'mfg
			plngCacheDuration = CMS_CacheDuration
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentPageTitle, contentMetaDescription, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = 3) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
			pstrGeneralSQL = "SELECT mfgID, mfgName From sfManufacturers Order By mfgName"
			pbytKeyIndex = 0
		Case 5	'Product Reviews
			plngCacheDuration = CMS_CacheDuration
			pbytKeyIndex = 4
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentContent, contentTitle, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = 5) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
		Case 6	'FAQ
			plngCacheDuration = CMS_CacheDuration
			pbytKeyIndex = 4
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentContent, contentTitle, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = 6) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
		Case Else
			plngCacheDuration = CMS_CacheDuration
			pbytKeyIndex = 4
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentContent, contentTitle, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = " & lngContentType & ") And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))" _
						   & " Order By contentSortOrder, contentTitle Asc"
	End Select

	If isValidCMSFile(strCMSFileLocation, plngCacheDuration) Then
		pblnValidData = loadXML(strCMSFileLocation)
	Else
		pblnValidData = False
	End If

	If Not pblnValidData Then
	
		Set	pdicItems = CreateObject("scripting.dictionary")
		
		set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 3			'adUseClient
			.Open pstrContentSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
			'Set .Connection = Nothing	'Disconnect the recordset
			
			If Not .EOF Then
				paryResult = .GetRows()
				.Close
				
				'Response.Write "<fieldset><legend>" & UBound(paryResult, 2) & " result(s) found in content for contentType of " & lngContentType & "</legend>SQL: " & pstrContentSQL & "</fieldset>"
				For i = 0 To UBound(paryResult, 2)
					pstrKey = "key" & CStr(paryResult(pbytKeyIndex,i))
					If Not pdicItems.Exists(pstrKey) Then pdicItems.Add pstrKey, CStr(paryResult(pbytKeyIndex,i))
					addItem paryResult(0,i), "cmsID=" & paryResult(4,i), paryResult(2,i), paryResult(3,i), paryResult(1,i), paryResult(4,i), ""
							'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref, count
					If cblnDebugCMS Then
						Response.Write "<fieldset><legend>getCMSData</legend>"
						Response.Write "contentReferenceID: " & paryResult(0,i) & "<BR>"
						Response.Write "contentPageName: " & paryResult(1,i) & "<BR>"
						Response.Write "contentPageTitle: " & paryResult(2,i) & "<BR>"
						Response.Write "contentMetaDescription: " & paryResult(3,i) & "<BR>"
						Response.Write "contentID: " & paryResult(4,i) & "<BR>"
						Response.Write "</fieldset>"
					End If
				Next
			Else
				'Response.Write "<fieldset><legend>No results found in content for contentType of " & lngContentType & "</legend>SQL: " & pstrContentSQL & "</fieldset>"
				.Close
			End If

			If Len(pstrGeneralSQL) > 0 Then
				.Open pstrGeneralSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
				'Set .Connection = Nothing	'Disconnect the recordset

				If Not .EOF Then
					paryResult = .GetRows()
					.Close
					
					For i = 0 To UBound(paryResult, 2)
						pstrKey = "key" & CStr(paryResult(0,i))
						If Not pdicItems.Exists(pstrKey) Then
							pdicItems.Add pstrKey, ""	'technically this shouldn't be necessary
							addItem paryResult(1,i), "mfgID=" & paryResult(0,i), "", "", "", "", ""
										'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref
						End If
					Next
				Else
					.Close
				End If
			End If
				
		End With	'pobjRS
		Set	pobjRS = Nothing
		Set	pdicItems = Nothing
		
		saveXML strCMSFileLocation

	End If	'pblnValidData
	
	getCMSData = True
	'Response.Write Server.HTMLEncode(XMLText)

End Function	'getCMSData

'******************************************************************************************************************************************************************

Function loadCMSData(byVal lngContentType)

Dim i
Dim paryResult
Dim pblnValidData
Dim pdicItems
Dim pobjRS
Dim pstrContentSQL
Dim pstrGeneralSQL
Dim pstrKey
Dim pbytKeyIndex

Dim pstrCMSFileLocation
Dim plngCacheDuration

	Select Case lngContentType
		Case 3	'mfg
			pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\contentType_" & lngContentType & ".xml"
			plngCacheDuration = CMS_CacheDuration
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentPageTitle, contentMetaDescription, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = 3) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
			pstrGeneralSQL = "SELECT mfgID, mfgName From sfManufacturers Order By mfgName"
			pbytKeyIndex = 0
		Case 5	'Product Reviews
			pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\contentType_" & lngContentType & ".xml"
			plngCacheDuration = CMS_CacheDuration
			pbytKeyIndex = 4
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentContent, contentTitle, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = 5) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
		Case 6	'FAQ
			pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\contentType_" & lngContentType & ".xml"
			plngCacheDuration = CMS_CacheDuration
			pbytKeyIndex = 4
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentContent, contentTitle, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = 6) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
		Case Else
			pstrCMSFileLocation = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSCache\contentType_" & lngContentType & ".xml"
			plngCacheDuration = CMS_CacheDuration
			pbytKeyIndex = 4
			pstrContentSQL = "SELECT contentReferenceID, contentPageName, contentContent, contentTitle, contentID, contentContentFilePath" _
						   & " FROM content Where (contentContentType = " & lngContentType & ") And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))" _
						   & " Order By contentSortOrder, contentTitle Asc"
	End Select
	
	If isValidCMSFile(pstrCMSFileLocation, plngCacheDuration) Then
		pblnValidData = loadXML(pstrCMSFileLocation)
	Else
		pblnValidData = False
	End If

	If Not pblnValidData Then
	
		Set	pdicItems = CreateObject("scripting.dictionary")
		
		set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 3			'adUseClient
			.Open pstrContentSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
			'Set .Connection = Nothing	'Disconnect the recordset
			
			If Not .EOF Then
				paryResult = .GetRows()
				.Close
				
				'Response.Write "<fieldset><legend>" & UBound(paryResult, 2) & " result(s) found in content for contentType of " & lngContentType & "</legend>SQL: " & pstrContentSQL & "</fieldset>"
				For i = 0 To UBound(paryResult, 2)
					pstrKey = "key" & CStr(paryResult(pbytKeyIndex,i))
					If Not pdicItems.Exists(pstrKey) Then pdicItems.Add pstrKey, CStr(paryResult(pbytKeyIndex,i))
					addItem paryResult(0,i), "cmsID=" & paryResult(4,i), paryResult(2,i), paryResult(3,i), paryResult(1,i), paryResult(4,i), ""
							'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref, count
					If cblnDebugCMS Then
						Response.Write "<fieldset><legend>loadCMSData</legend>"
						Response.Write "contentReferenceID: " & paryResult(0,i) & "<br />"
						Response.Write "contentPageName: " & paryResult(1,i) & "<br />"
						Response.Write "contentPageTitle: " & paryResult(2,i) & "<br />"
						Response.Write "contentMetaDescription: " & paryResult(3,i) & "<br />"
						Response.Write "contentID: " & paryResult(4,i) & "<br />"
						Response.Write "</fieldset>"
					End If
				Next
			Else
				'Response.Write "<fieldset><legend>No results found in content for contentType of " & lngContentType & "</legend>SQL: " & pstrContentSQL & "</fieldset>"
				.Close
			End If

			If Len(pstrGeneralSQL) > 0 Then
				.Open pstrGeneralSQL, cnn, 0,1		'adOpenForwardOnly,adLockReadOnly
				'Set .Connection = Nothing	'Disconnect the recordset

				If Not .EOF Then
					paryResult = .GetRows()
					.Close
					
					For i = 0 To UBound(paryResult, 2)
						pstrKey = "key" & CStr(paryResult(0,i))
						If Not pdicItems.Exists(pstrKey) Then
							pdicItems.Add pstrKey, ""	'technically this shouldn't be necessary
							addItem paryResult(1,i), "mfgID=" & paryResult(0,i), "", "", "", "", ""
										'pstrName, pstrValue, pstrtext, pstrTitle, pstrhref
						End If
					Next
				Else
					.Close
				End If
			End If
				
		End With	'pobjRS
		Set	pobjRS = Nothing
		Set	pdicItems = Nothing
		
		If plngCacheDuration > 0 Then saveXML pstrCMSFileLocation

	End If	'pblnValidData
	
	loadCMSData = True
	'Response.Write Server.HTMLEncode(XMLText)

End Function	'loadCMSData

'***************************************************************************************************************************************************************

Public Function LoadProducts(byVal lngContentType, byVal lngContentID)

Dim pobjRS
Dim pstrSQL

'On Error Resume Next

	If Len(plngPageNumber) = 0 Then plngPageNumber = 1
	If plngPageNumber = 0 Then plngPageNumber = 1
	
	'Determine number of products which meet this content type
	'Sources are products table and content table: some duplication may occur
	Select Case lngContentType
		Case 3	'mfg
			pstrSQL = "SELECT contentReferenceID, contentPageName, contentPageTitle, contentMetaDescription, contentID" _
						   & " FROM content Where (contentContentType = 3) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
			pstrSQL = "SELECT prodID, prodName, prodShortDescription, prodImageSmallPath From sfProducts Where prodEnabledIsActive=1 And prodManufacturerId=" & lngContentID & " Order By prodName"
		Case enContentType_Breeds	'mfg
			pstrSQL = "SELECT contentReferenceID, contentPageName, contentPageTitle, contentMetaDescription, contentID" _
						   & " FROM content Where (contentContentType = 3) And ((contentApprovedForDisplay = 1) Or (contentApprovedForDisplay = -1))"
			pstrSQL = "SELECT prodID, prodName, prodShortDescription, prodImageSmallPath From sfProducts Where prodEnabledIsActive=1 And prodManufacturerId=" & lngContentID & " Order By prodName"
		Case Else
			LoadProducts = False
			Exit Function
	End Select

	Set	pobjRS = CreateObject("adodb.recordset")
	With pobjRS
        .CursorLocation = 3 'adUseClient

		If Len(plngPageSize) > 0 Then 
			.CacheSize = plngPageSize
			.PageSize = plngPageSize
		End If

		On Error Resume Next
		
		.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		If Err.number <> 0 Then
			Response.Write "<fieldset><legend>Error loading products</legend>"
			Response.Write "Error " & err.number & ": " & err.Description & "<br />"
			Response.Write "SQL: " & pstrSQL & "<br />"
			Response.Write "</fieldset>"
			err.Clear
		ElseIf Not .EOF Then
			plngRecordCount = .RecordCount
			plngPageCount = .PageCount
			If cInt(plngPageNumber) > cInt(plngPageCount) Then plngPageNumber = plngPageCount
			.AbsolutePosition = (plngPageNumber - 1) * .PageSize + 1
		End If
		.Close

	End With	'pobjRS
	Set pobjRS = Nothing

	If True Then
		Response.Write "<fieldset><legend>LoadProducts Results</legend>"
		Response.Write "lngContentType: " & lngContentType & "<br />"
		Response.Write "plngRecordCount: " & plngRecordCount & "<br />"
		Response.Write "plngPageNumber: " & plngPageNumber & "<br />"
		Response.Write "plngPageSize: " & plngPageSize & "<br />"
		Response.Write "</fieldset>"
		err.Clear
		Response.Flush
	End If

    LoadProducts = (plngRecordCount > 0)

End Function    'LoadProducts

'***********************************************************************************************

Public Function WriteXSL(byVal lngContentType)

Dim objXSL
Dim pstrXSLFilePath
Dim strOutput

	If Len(pstrTemplatePath) = 0 Then
		Select Case lngContentType
			Case 3	'mfg
				pstrXSLFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSTemplates\links.xsl"
			Case 6	'FAQ
				pstrXSLFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSTemplates\faq.xsl"
			Case Else
				pstrXSLFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "\CMSTemplates\links.xsl"
		End Select
	Else
		pstrXSLFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & pstrTemplatePath
	End If

	' Load the XSL from the XSL file
	set objXSL = CreateObject("MSXML2.DOMDocument")
	objXSL.async = false
	'objXSL.preserveWhiteSpace = True
	
	'On Error Resume Next
	If objXSL.Load(pstrXSLFilePath) Then
		If False Then
			Response.Write "<fieldset><legend>WriteXSL - XSL</legend>"
			Call TestWriteXML(objXSL)
			Response.Write "</fieldset>"
		End If
		If isObject(p_objxmlDoc) Then strOutput = p_objxmlDoc.transformNode(objXSL)
		'More efficient for longer documents
		'p_objxmlDoc.transformNodeToObject objXSL, Response
	Else
		Dim myErr
		Set myErr = objXSL.parseError
		strOutput = "Error Loading XSL document " & pstrXSLFilePath & ": " & myErr.reason
	End If
	Set objXSL = Nothing
	
	If Err.number <> 0 Then
		strOutput = "Error " & err.number & ": " & err.Description & "<br />"
		Err.Clear
	End If
	
	'strOutput = Replace(strOutput,"&amp;nbsp;","&nbsp;")
	
	WriteXSL = strOutput

End Function	'WriteXSL

'******************************************************************************************************************************************************************

End Class   'clsXMLCMS
%>
