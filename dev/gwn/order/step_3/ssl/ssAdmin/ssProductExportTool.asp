<%Option Explicit
'********************************************************************************
'*   Product Pricing Tool For StoreFront 6.0
'*   Release Version:	2.00.001
'*   Release Date:		January 1, 2006
'*   Revision Date:		January 1, 2006
'*
'*   Release Notes:
'*
'*   2.00.001 - January 1, 2006
'*	 ' Initial Release
'*
'*   The contents of this file are protected by United States copyright laws
'*   as an unpublished work. No part of this file may be used or disclosed
'*   without the express written permission of Sandshot Software.
'*
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.
'********************************************************************************

Response.Buffer = True
Server.ScriptTimeout = 900			'in seconds. Adjust for large databases or if some products have a lot of attributes. Server Default is usually 90 seconds

'Add-on Settings
'pstrssAddonCode = "ProductExportSF6v2"
mstrssAddonVersion = "2.00.001"

'***********************************************************************************************

Class clsProductExport
'Assumptions:
'   cnn: defines a previously opened connection to the database

'class variables
Private cstrDelimeter
Private pstrMessage
Private prsProducts
Private pblnError
Private cblnDebug
Private xmlDoc
Private pblnPreventDuplicates
Private pblnUseAttributes
Private pblnUseAllAttributes
Private pblnShowCost
Private plngProductCount

'***********************************************************************************************

Private Sub class_Initialize()
    
	pblnUseAttributes = True	'True	False
	pblnUseAllAttributes = True	'True	False
	pblnPreventDuplicates = True
	
    cstrDelimeter  = ";"
    cblnDebug = False	'True	False
    
End Sub

Private Sub class_Terminate()
    On Error Resume Next
	Call ReleaseObject(prsProducts)
	Call ReleaseObject(xmlDoc)
End Sub

'***********************************************************************************************

Public Property Get Message()
    Message = pstrMessage
End Property

Public Property Get UseAttributes()
    UseAttributes = pblnUseAttributes
End Property

Public Property Let UseAttributes(byVal vntValue)
    pblnUseAttributes = vntValue
End Property

Public Property Let UseAllAttributes(byVal vntValue)
    pblnUseAllAttributes = vntValue
End Property

Public Property Let PreventDuplicates(byVal vntValue)
    pblnPreventDuplicates = vntValue
End Property

Public Property Get rsProducts()
    Set rsProducts = prsProducts
End Property

Public Property Get ProductCount()
    ProductCount = plngProductCount
End Property

Public Property Get xmlDSO()

Dim node

	Set node = xmlDoc.getElementsByTagName("products")
    xmlDSO = node.item(0).xml
    
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

Public Function Load()

dim pstrSQL
dim p_strWhere
dim i

'On Error Resume Next

	Dim pstrCategoryJoin
	
	pstrCategoryJoin = "RIGHT" 'All categories
	pstrCategoryJoin = "INNER" 'One category

	If len(mlngAbsolutePage) = 0 Then mlngAbsolutePage = 1
	
	set	prsProducts = server.CreateObject("adodb.recordset")
	With prsProducts
        .CursorLocation = 3 'adUseClient

		pstrSQL = "SELECT Products.uid" _
				& " FROM ProductCategory " & pstrCategoryJoin & " JOIN Products ON ProductCategory.ProductID = Products.uid" _
				& " " & mstrSQLWhere _
				& " GROUP BY Products.uid, Products.Code, Products.Name, Products.IsActive, Products.Name, Products.Cost, Products.Price, Products.IsOnSale, Products.SalePrice, Products.ManufacturerId, Products.VendorId" _
				& mstrsqlHaving _
				& " ORDER BY Products.Code"

		If (Len(mbytCategoryFilter) > 0 And cblnSF5AE) OR mblnShowUnassignedProducts Then
			If mblnShowUnassignedProducts Then
				'Distinct modifier removed because it fails with an order by clause
				'pstrSQL = "Select distinct sfProducts.prodID" _
				pstrSQL = "Select sfProducts.prodID" _
						& " FROM (sfProducts LEFT JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) LEFT JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " _
						& mstrSQLWhere
			Else
				'Distinct modifier removed because it fails with an order by clause
				'pstrSQL = "Select distinct sfProducts.prodID" _
				pstrSQL = "Select sfProducts.prodID" _
						& " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID " _
						& mstrSQLWhere
			End If
	    Else
	        pstrSQL = "Select sfProducts.prodID" _
					& " FROM sfProducts " _
					& mstrsqlWhere
	    End If

		If cblnDebug Then debugprint "pstrSQL",pstrSQL
		If cblnDebug Then Response.Flush	  
		  
		'On Error Resume Next
		'.Open pstrSQL, cnn, 1,1	'adOpenKeySet,adLockReadOnly
		.Open pstrSQL, cnn, 3,1	'adOpenKeySet,adLockReadOnly
		
		If Err.number <> 0 Then
			Response.Write "<font color=red>Error in Load: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
			Response.Write "<font color=red>Error in Load: pstrSQL = " & pstrSQL & "</font><br />" & vbcrlf
			Call ShowStoreFrontVersion
			Response.Flush
			Err.Clear
			Load = False
			Exit Function
		End If
		
		If cblnDebug Then debugprint ".EOF",.EOF
		If .EOF Then
			mlngPageCount = 0
			plngProductCount = 0
		Else
			mlngPageCount = .PageCount
			plngProductCount = .RecordCount
		End If
		If cblnDebug Then debugprint "mlngPageCount", mlngPageCount
		If cInt(mlngAbsolutePage) > cInt(mlngPageCount) Then mlngAbsolutePage = mlngPageCount
		
	End With

	If prsProducts.EOF Then
		Load = False
	Else
		Call ssCreateProductXML(prsProducts.getString(2,,"",","))
		Load = True
	End If
	    
End Function    'Load

'***********************************************************************************************

Function ssCreateProductXML(byVal strUIDs)

Dim fieldCounter
Dim paryProductUIDs
Dim pblnAddAttribute
Dim pblnIgnore
Dim pdicVendors
Dim pdicManufacturers
Dim plngProductCounter
Dim pobjRS
Dim pobjCmd
Dim pobjRSCategory
Dim pstrCategoryName
Dim pstrFieldName
Dim pstrFieldValue
Dim pstrKey
Dim pstrPrevID
Dim pstrPrevAttrID
Dim pstrSQL
Dim pxmlExtraItem
Dim pxmlExtras
Dim xmlRoot
Dim xmlNode
Dim xmlProduct
Dim xmlAttributes
Dim xmlAttributeCategory
Dim xmlAttributeDetail
Dim pdicCategories
Dim xmlProductCategories
Dim xmlProductCategory
Dim XPath

	Set pdicCategories = Server.CreateObject("Scripting.Dictionary")
	Set pdicVendors = Server.CreateObject("Scripting.Dictionary")
	Set pdicManufacturers = Server.CreateObject("Scripting.Dictionary")
	
	Set xmlDoc = server.CreateObject("MSXML2.DOMDocument.3.0")
	' Create processing instruction and document root
    Set xmlNode = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'")
    Set xmlNode = xmlDoc.insertBefore(xmlNode, xmlDoc.childNodes.Item(0))
   
	' Create document root
    Set xmlRoot = xmlDoc.createElement("products")
    Set xmlDoc.documentElement = xmlRoot
    xmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"

	If Len(strUIDs) > 0 Then
		If Right(strUIDs, 1) = "," Then strUIDs = Left(strUIDs, Len(strUIDs) - 1)
	End If
	paryProductUIDs = Split(strUIDs, ",")
	
	pstrSQL = "SELECT sfProducts.*, sfAttributes.*, sfAttributeDetail.*, sfManufacturers.mfgName, sfVendors.vendName" _
			& " FROM sfVendors INNER JOIN (((sfManufacturers INNER JOIN sfProducts ON sfManufacturers.mfgID = sfProducts.prodManufacturerId) LEFT JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) LEFT JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId) ON sfVendors.vendID = sfProducts.prodVendorId" _
			& " WHERE sfProducts.prodID=?" _
			& " ORDER BY sfProducts.prodID, sfAttributes.attrName, sfAttributeDetail.attrdtName"
	
	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		.Commandtext = pstrSQL
		Set .ActiveConnection = cnn
		.Parameters.Append .CreateParameter("prodUID", adVarChar, adParamInput, 50, 0)
	End With	'pobjCmd
	pstrSQL = ""

	For plngProductCounter = 0 To UBound(paryProductUIDs)
		pstrPrevID = paryProductUIDs(plngProductCounter)

		pobjCmd.Parameters("prodUID").Value = pstrPrevID
		Set pobjRS = pobjCmd.Execute
		With pobjRS
		
			'Create Product Node
			Set xmlProduct = xmlDoc.createElement("product")
			xmlRoot.appendChild xmlProduct
			xmlProduct.setAttribute "uid", pstrPrevID

			'Now we need to add the categories
			Set xmlProductCategories = xmlDoc.createElement("categories")
			xmlProduct.appendChild xmlProductCategories
				
			mlngCategoryFilter = LoadRequestValue("CategoryFilter")
			If cblnSF5AE Then
				'If Len(mlngCategoryFilter) > 0 Then
					'pstrSQL = "Select CategoryID From ProductCategory Where ProductID=" & wrapSQLValue(pstrPrevID, False, enDatatype_number) & " And CategoryID In (" & childCategories(mlngCategoryFilter) & ")"
					pstrSQL = "Select subcatCategoryId As CategoryID From sfSubCatDetail Where ProdID=" & wrapSQLValue(pstrPrevID, False, enDatatype_string)
				'End If

				Set pobjRSCategory = GetRS(pstrSQL)
				With pobjRSCategory
					Do While Not .EOF
						If .Fields("CategoryID").Value <> 1 Or True Then
							Set xmlProductCategory = xmlDoc.createElement("category")
							xmlProductCategories.appendChild xmlProductCategory
							xmlProductCategory.setAttribute "CategoryID", .Fields("CategoryID").Value
							
							If pdicCategories.Exists(.Fields("CategoryID").Value) Then
								pstrCategoryName = pdicCategories(.Fields("CategoryID").Value)
							Else
								pstrCategoryName = getCategoryName(.Fields("CategoryID").Value, ">")
							End If
							Call addCDATA(xmlDoc, xmlProductCategory, "categoryName", pstrCategoryName)
							
						End If	'.Fields("CategoryID").Value <> 1
					
						.MoveNext
					Loop
					.Close
				End With	'pobjRSCategory
				Set pobjRSCategory = Nothing
			Else
				If Not .EOF Then
					Set xmlProductCategory = xmlDoc.createElement("category")
					xmlProductCategories.appendChild xmlProductCategory
					xmlProductCategory.setAttribute "CategoryID", .Fields("prodCategoryId").Value
					pstrCategoryName = getCategoryName(.Fields("prodCategoryId").Value, ">")
					Call addCDATA(xmlDoc, xmlProductCategory, "categoryName", pstrCategoryName)
				End If
			End If

			'add the root order elements
			For fieldCounter = 1 To .Fields.Count
				pstrFieldName = Trim(.Fields(fieldCounter-1).Name & "")
				pstrFieldValue = getRSFieldValue_Unknown(.Fields(fieldCounter-1))
				
				'remove carriage returns
				pstrFieldValue = Replace(pstrFieldValue, vbcrlf, "")

				pblnIgnore = False
				Select Case pstrFieldName
					Case "attrID", "attrdtID", "attrName", "attrdtName", "attrdtPrice", "attrdtPriceTyp", "AttributeOrder"
						pblnIgnore = True
					Case "prodLink"
						If Left(pstrFieldValue, 3) = "../" Then pstrFieldValue = mstrBaseHRef & Replace(pstrFieldValue, "../", "", 1, 1)
						If LCase(Left(pstrFieldValue, 5)) <> "http:" Then pstrFieldValue = mstrBaseHRef & "detail.asp?Product_Id=" & Replace(pstrFieldValue, " ", "+")
					Case "prodImageSmallPath", "prodImageLargePath"
						If Left(LCase(pstrFieldValue), 4) <> "http" Then pstrFieldValue = mstrBaseHRef & pstrFieldValue
						If Left(pstrFieldValue, 1) = "/" Then pstrFieldValue = mstrBaseHRef & Replace(pstrFieldValue, "/", "", 1, 1)
					Case "someFieldName"
						If Len(pstrFieldValue) = 0 Then pstrFieldValue = 0
					Case Else
						'do nothing, already covered
				End Select
				If Not pblnIgnore Then Call addNode(xmlDoc, xmlProduct, pstrFieldName, pstrFieldValue)
			Next 'fieldCounter

			'now for the attributes
			Set xmlAttributes = xmlDoc.createElement("attributes")
			xmlProduct.appendChild xmlAttributes
			Do While Not .EOF
				pblnAddAttribute = Not isNull(.Fields("attrID").Value)
				If CBool(pstrPrevAttrID <> Trim(.Fields("attrID").Value & "")) And pblnAddAttribute Then
					pstrPrevAttrID = CStr(.Fields("attrID").Value)
					Set xmlAttributeCategory = xmlDoc.createElement("attributeCategory")
					xmlAttributes.appendChild xmlAttributeCategory
					xmlAttributeCategory.setAttribute "uid", pstrPrevAttrID
					Call addNode(xmlDoc, xmlAttributeCategory, "name", Trim(.Fields("attrName").Value & ""))
				End If
			
				If pblnAddAttribute And Not isNull(.Fields("attrdtID").Value) Then
					'Create Attribute Detail Node
					Set xmlAttributeDetail = xmlDoc.createElement("attribute")
					xmlAttributeCategory.appendChild xmlAttributeDetail
					xmlAttributeDetail.setAttribute "uid", Trim(.Fields("attrdtID").Value & "")
					
					Call addNode(xmlDoc, xmlAttributeDetail, "name", Trim(.Fields("attrdtName").Value & ""))
					Select Case Trim(.Fields("attrdtType").Value & "")		
						Case "1": pstrFieldValue = Trim(.Fields("attrdtPrice").Value & "")
						Case "2": pstrFieldValue = Trim(-1 * .Fields("attrdtPrice").Value & "")
						Case Else: pstrFieldValue = "0"
					End Select
					Call addNode(xmlDoc, xmlAttributeDetail, "price", pstrFieldValue)
				End If

				If Not pdicVendors.Exists(.Fields("prodVendorId").Value) Then pdicVendors.Add .Fields("prodVendorId").Value, .Fields("mfgName").Value
				If Not pdicManufacturers.Exists(.Fields("prodManufacturerId").Value) Then pdicManufacturers.Add .Fields("prodManufacturerId").Value, .Fields("vendName").Value
				.MoveNext
			Loop
			.Close
		End With	'	pobjRS	
	Next 'plngProductCounter
	Set pobjRS = Nothing
	Set pobjCmd = Nothing
	
	'add the Mfg information
	'Create Mfgs Node
	Set pxmlExtras = xmlDoc.createElement("Manufacturers")
	xmlRoot.appendChild pxmlExtras
	For Each pstrKey in pdicManufacturers
		Set pxmlExtraItem = xmlDoc.createElement("Manufacturer")
		pxmlExtras.appendChild pxmlExtraItem
		pxmlExtraItem.setAttribute "uid", pstrKey
		Call addNode(xmlDoc, pxmlExtraItem, "name", pdicManufacturers(pstrKey))
	Next	'pdicManufacturers
	Set pxmlExtras = Nothing

	Set pxmlExtras = xmlDoc.createElement("Vendors")
	xmlRoot.appendChild pxmlExtras
	For Each pstrKey in pdicVendors
		Set pxmlExtraItem = xmlDoc.createElement("Vendor")
		pxmlExtras.appendChild pxmlExtraItem
		pxmlExtraItem.setAttribute "uid", pstrKey
		Call addNode(xmlDoc, pxmlExtraItem, "uid", pstrKey)
		Call addNode(xmlDoc, pxmlExtraItem, "name", pdicVendors(pstrKey))
	Next	'pdicVendors
	Set pxmlExtras = Nothing

	ssCreateProductXML = xmlDoc.xml
	
	'xmlDoc.preserveWhiteSpace = False
	'Response.Write "<fieldset><legend>ssCreateOrderXML</legend><textarea rows=80 cols=120>" & vbcrlf & xmlDoc.xml & vbcrlf & "</textarea></fieldset>"

	Set pdicCategories = Nothing
	Set pdicVendors = Nothing
	Set pdicManufacturers = Nothing

End Function	'ssCreateProductXML

'***************************************************************************************************************************************************************

Function transformData(byVal strXSLFilePath)

Dim objXSL
Dim strOutput

	If Not isObject(xmlDoc) Then Exit Function
	
	' Load the XSL from the XSL file
	set objXSL = Server.CreateObject("MSXML2.DOMDocument.3.0")
	objXSL.async = false
	objXSL.preserveWhiteSpace = True
	'debugprint "strXSLFilePath", strXSLFilePath
	If objXSL.Load(strXSLFilePath) Then
		strOutput = xmlDoc.transformNode(objXSL)
	Else
		Dim myErr
		Set myErr = objXSL.parseError
		strOutput = "Error loading " & strXSLFilePath & ". Error " & myErr.errorCode & ": " & myErr.reason
	End If
	Set objXSL = Nothing
	
	strOutput = Replace(strOutput,"&amp;nbsp;","&nbsp;")
	transformData = strOutput

End Function	'transformData

'***************************************************************************************************************************************************************

End Class   'clsProductExport

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="Common/ssProduct_CommonFilter.asp"-->
<!--#include file="Common/ssProducts_Common.asp"-->
<!--#include file="AdminHeader.asp"-->
<%

'******************************************************************************************************************************************************************
'
'	Begin Main Page
'
'******************************************************************************************************************************************************************

'page variables
Dim i
Dim cstrDefaultProductExportFilename
Dim maryTemp
Dim maryExportTemplates
Dim mblnDisplayOutput
Dim mblnShowFilter
Dim mclsProductExport
Dim mlngAbsolutePage
Dim mlngMaxRecords
Dim mlngPageCount
Dim mstrAction
Dim mstrXSLFilePath
Dim mblnSuppressOutput
Dim mstrTempOutput
Dim mstrTempFileName

	'mstrPageTitle = CheckForUpdatedVersion(pstrssAddonCode, pstrssAddonVersion)
	mstrPageTitle = "Product Export Tool"

	mblnShowFilter = (lCase(trim(Request.Form("blnShowFilter"))) = "true")
	
	Call getFileNamesInFolder(ssAdminPath & "exportTemplates\ProductExportToolTemplates\", ".xsl", maryExportTemplates)
	mlngAbsolutePage = LoadRequestValue("AbsolutePage")
	mblnSuppressOutput = CBool(LoadRequestValue("chkSuppressOutput") = "1")

	mblnDisplayOutput = True
	mstrAction = LoadRequestValue("Action")
	If Len(LoadRequestValue("Action")) > 0 Then
		maryTemp = Split(mstrAction, "|")
		mstrAction = maryTemp(0)
		If UBound(maryTemp) > 0 Then mstrXSLFilePath = maryTemp(1)
	End If
	
	'Call LoadProductFilterFromRequest
    Call LoadFilter

    Set mclsProductExport = New clsProductExport
	cstrDefaultProductExportFilename = getAddonConfigurationSetting("ProductExport", "cstrDefaultProductExportFilename")
    mclsProductExport.UseAttributes = getAddonConfigurationSetting("ProductExport", "cblnDefaultProductExportUseAttributes")
    mclsProductExport.UseAllAttributes = getAddonConfigurationSetting("ProductExport", "cblnDefaultProductExportUseAllAttributes")
    mclsProductExport.PreventDuplicates = getAddonConfigurationSetting("ProductExport", "cblnDefaultProductExportPreventDuplicates")

    Select Case mstrAction
        Case "downloadItems"
			If Len(cstrDefaultProductExportFilename) = 0 Then
				mstrTempFileName = Left(mstrXSLFilePath, Len(mstrXSLFilePath) - 4) & ".csv"
			Else
				mstrTempFileName = cstrDefaultProductExportFilename
			End If
			mstrXSLFilePath = Server.MapPath("exportTemplates/ProductExportToolTemplates/" & mstrXSLFilePath)
			
			If mblnSuppressOutput Then
				If Response.Buffer Then Response.Flush
				Response.Write "<h4>Download template " & mstrXSLFilePath & "</h4>"
				Response.Write "Loading Product Data . . .<br />"
				mclsProductExport.Load
				Response.Write "Product Data Loaded. Processing . . .<br />"
				mstrTempOutput = mclsProductExport.transformData(mstrXSLFilePath)
				Response.Write "Processing Complete. Saving to file . . .<br />"
				If writeToFile("[webroot]\ssl\admin\ssAdmin\ssExportedFiles\" & mstrTempFileName, mstrTempOutput) Then
					Response.Write "Your data file is ready to download <a href=""ssExportedFiles\" & mstrTempFileName & """>here</a>."
				End If
				mblnDisplayOutput = False
			Else
				mclsProductExport.Load
				mstrTempOutput = mclsProductExport.transformData(mstrXSLFilePath)
				Response.ContentType = "application/octet-stream"
				Response.AddHeader "Content-Disposition", "attachment; filename=""" & mstrTempFileName & """"
				Response.Write mstrTempOutput
				mblnDisplayOutput = False
			End If	'mblnSuppressOutput

		Case "Filter"
			mclsProductExport.Load
			mblnShowFilter = False
        Case "printItems"
			mstrXSLFilePath = Server.MapPath("exportTemplates/ProductExportToolTemplates/" & mstrXSLFilePath)
			mclsProductExport.Load
			Response.Write mclsProductExport.transformData(mstrXSLFilePath)
			Response.Write "<OBJECT ID=WebBrowser1 WIDTH=0 HEIGHT=0 CLASSID='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2'></OBJECT>"
			Response.Write "<script language=javascript>" _
						   & "document.all('WebBrowser1').ExecWB(6, 2);window.close();</script>"
			mblnDisplayOutput = False
        Case "viewItems"
			mstrXSLFilePath = Server.MapPath("exportTemplates/ProductExportToolTemplates/" & mstrXSLFilePath)
			mclsProductExport.Load
			Response.Write mclsProductExport.transformData(mstrXSLFilePath)
			mblnDisplayOutput = False
		Case Else
			'Do Nothing
    End Select	'mstrAction
    
    If Not mblnDisplayOutput Then
		Call ReleaseObject(mclsProductExport)
		Call ReleaseObject(cnn)
		Response.Flush
		Response.End
    End If

	Call WriteHeader("body_onload();",True)
%>
<script LANGUAGE="JavaScript" SRC="SSLibrary/ssFormValidation.js"></script>
<script LANGUAGE=javascript>
<!--

var theDataForm;

function body_onload()
{
	theDataForm = document.frmData;
}

function checkAll_custom(strType, blnCheck)
{
	var plngCount;
	var i;
	var strCheckboxName;

	switch (strType)
	{
		case 'update':
			checkAll(document.frmData.prodID, blnCheck)
			checkAll(document.frmData.attrdtID, blnCheck)
			break;
		case 'prodSaleIsActive':
			plngCount = document.frmData.prodID.length;
			if (document.frmData.prodID.checked==undefined)
			{
				for (i=0; i < plngCount;i++)
				{
				strCheckboxName = strType + document.frmData.prodID[i].value;
				checkAll(document.all(strCheckboxName), blnCheck)
				}
			}else{
				strCheckboxName = strType + document.frmData.prodID.value;
				checkAll(document.all(strCheckboxName), blnCheck)
			}
			break;
		case 'IsActive':
			plngCount = document.frmData.prodID.length;
			if (document.frmData.prodID.checked==undefined)
			{
				for (i=0; i < plngCount;i++)
				{
				strCheckboxName = strType + document.frmData.prodID[i].value;
				checkAll(document.all(strCheckboxName), blnCheck)
				}
			}else{
				strCheckboxName = strType + document.frmData.prodID.value;
				checkAll(document.all(strCheckboxName), blnCheck)
			}
			break;
	}
	return true;

	if (document.frmData.attrdtID != undefined)
	{
		plngCount = document.frmData.attrdtID.length;
		if (document.frmData.attrdtID.checked==undefined)
		{
			for (i=0; i < plngCount;i++)
			{
			document.frmData.attrdtID[i].checked = blnCheck;
			}
		}else{
			document.frmData.attrdtID.checked = blnCheck;
		}
	}
	
}

function downloadItems()
{
	var Template = templateSelected();
	
	if (Template != false)
	{
		if (itemSelected())
		{
			submitForm('downloadItems' + '|' + Template, "downloadItems")
		}
	}
	return false;
}

function itemSelected()
{
	return true;
	if (! anyChecked(theDataForm.chkOrderUID))
	{
		alert("Please select at least one order to view.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	
	return true;
}

function printItems(strTemplate)
{
	var Template = templateSelected();
	
	if (Template != false)
	{
		if (itemSelected())
		{
			submitForm('printItems' + '|' + Template, "printItems")
		}
	}
	return false;
}

function viewItems(strTemplate)
{
	var Template = templateSelected();
	
	if (Template != false)
	{
		if (itemSelected())
		{
			submitForm('viewItems' + '|' + Template, "viewItems")
		}
	}
	return false;
}

function submitForm(strAction, strTarget)
{
	var prevAction = theDataForm.Action.value;

	theDataForm.Action.value = strAction;
	theDataForm.target = strTarget;
	theDataForm.submit();
	theDataForm.Action.value = prevAction;
	theDataForm.target='';
	return false;
}

function templateSelected()
{
	if (theDataForm.ExportTemplates.selectedIndex == 0)
	{
		alert("Please select a template.");
		theDataForm.ExportTemplates.focus();
		return false;
	}
	return theDataForm.ExportTemplates.options[theDataForm.ExportTemplates.selectedIndex].value;
}

//-->
</script>
<% If Len(mstrAction) > 0 Then %>
<XML ID="dso">
  <% If mclsProductExport.ProductCount > 0 Then %>
  <%= mclsProductExport.xmlDSO %>
  <% End If %>
</XML>
<% End If %>
<center>

<form action="ssProductExportTool.asp" id="frmData" name="frmData" method="post">
<input type="hidden" id="Action" name="Action" value="Update">
<input type="hidden" id="blnShowFilter" name="blnShowFilter" value="">

<table border=0 cellPadding=5 cellSpacing=1 width="100%" ID="tblMain">
  <tr>
    <th align='right'>
		<span class="pagetitle2"><%= mstrPageTitle %></span>
		&nbsp;<img src="images/properties.gif" onclick="openProperties('ProductExport')" title="Configure Product Manager Export Tool Options">
		&nbsp;<input class="butn" type=button value="?" onclick="OpenHelp('ssHelpFiles/ProductManagerExportModule/help_PMExportModule.htm')" id="btnhel" name=btnHelp title="Release Version <%= mstrssAddonVersion %>"><br />
		<a href="#"><div id="divFilter" onclick="return DisplayFilter();" onMouseOver="return DisplayTitle(this);" onMouseOut="ClearTitle();" title="Hide Filter">Hide Filter</div></a>
	</th>
  </tr>
  <tr>
    <td align="right" valign=middle>
		<label for="chkSuppressOutput">Suppress detailed product output</label>&nbsp;
		<input type="checkbox" name="chkSuppressOutput" id="chkSuppressOutput" value="1" <%= isChecked(mblnSuppressOutput) %>>&nbsp;
		<select name="ExportTemplates" ID="ExportTemplates">
		  <option value="" selected>Select a Template</option>
		<% If isArray(maryExportTemplates) Then
		   For i = 0 To UBound(maryExportTemplates) %>
		  <option value="<%= maryExportTemplates(i) %>"><%= Replace(maryExportTemplates(i), ".xsl", "") %></option>
		<% 
		   Next 'i 
		   End If
		%>
		</select>
      <input class="butn" id="btnView" name="btnView" type=image src="images/preview.gif" value="View" onclick="viewItems(); return false;" title="View selected items">&nbsp;&nbsp;
      <input class="butn" id="btnPrint" name="btnPrint" type=image src="images/print.gif" value="Print" onclick="printItems(); return false;" title="Print selected items">&nbsp;&nbsp;
      <input class="butn" id="btnSave" name="btnSave" type=image src="images/save.gif" value="Save" onclick="return downloadItems(); return false;" title="Download selected items using the selected template">
      </td>
  </tr>
</table>

<% Call WriteProductFilter(False) %>
<% If Len(mstrAction) = 0 Then %>
<h3>Please select a filter</h3>
<% Else %>
<table class="tbl" width="100%" cellpadding="3" cellspacing="0" border="1" xrules="none" id="tblSummary">
  <% If mclsProductExport.ProductCount > 0 Then %>
  <% If mblnSuppressOutput Then
		Response.Write "<h3>Product List Suppressed</h3>"
		Response.Write "<h3>" & mclsProductExport.ProductCount & " products meet your filter criteria</h3>"
	 Else
		Response.Write mclsProductExport.transformData(Server.MapPath("ProductExportTool_Support/exportDisplay.xsl"))
	 End If
  %>
  <% Else %>
  <tr class="tblhdr">
	<th align="center" colspan="1">No Products meet your filter criteria</th>
  </tr>
  <% End If	'Not mclsProductExport.rsProducts.EOF %>
</table>

<script LANGUAGE="javascript">
<!--

function makeDirty(lngID, blnAttribute)
{

	var plngCount;
	var i;

	if (blnAttribute)
	{
		plngCount = document.frmData.attrdtID.length;
		if (document.frmData.attrdtID.checked==undefined)
		{
			for (i=0; i < plngCount;i++)
			{
				if (document.frmData.attrdtID[i].value == lngID)
				{
				document.frmData.attrdtID[i].checked = true;
				return true
				}
			}
		}
	}else{
		plngCount = document.frmData.prodID.length;
		if (plngCount!=undefined)
		{
			for (i=0; i < plngCount;i++)
			{
				if (document.frmData.prodID[i].value == lngID)
				{
				document.frmData.prodID[i].checked = true;
				return true
				}
			}
		}else{
			document.frmData.prodID.checked = true;		
		}
	}

}

function DisplaySection(strSection)
{
var arySections = new Array("Cost","Regular","Sale","Attribute");

 for (var i=0; i < arySections.length;i++)
 {
//alert(i + ": " strSection + " - " + (arySections[i] == strSection));
	if (arySections[i] == strSection)
	{
		document.all("tbl" + arySections[i]).style.display = "";
		document.all("td" + arySections[i]).className = "hdrSelected";
	}else{
		document.all("tbl" + arySections[i]).style.display = "none";
		document.all("td" + arySections[i]).className = "hdrNonSelected";
	}
 }	
 
return(false);
}

<% If Not mblnShowFilter then Response.Write "DisplayFilter();" & vbcrlf %>
//-->
</script>

<% End If	'Len(mstrAction) = 0 %>
</form>

</center>
<!--#include file="AdminFooter.asp"-->
</BODY>
</HTML>
<%
	Call ReleaseObject(mclsProductExport)
	Call ReleaseObject(cnn)
	If Response.Buffer Then Response.Flush
%>