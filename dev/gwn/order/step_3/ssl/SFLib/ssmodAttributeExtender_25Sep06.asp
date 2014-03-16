<!--#include file="ssmodAttributeExtender_legacy.asp"-->
<%
'********************************************************************************
'*   Attribute Extender for StoreFront 5.0										*
'*   Release Version   1.01.005													*
'*   Release Date      July 21, 2002											*
'*   Updated on	       January 22, 2004											*
'*																				*
'*   Release Notes:                                                             *
'*                                                                              *
'*   Version 1.01.005 (January 22, 2004)		                                *
'*	 - Resolved custom version differences with comma character text based issue*
'*	 - Note: incAE.asp must be modified to implement the above correction		*
'*                                                                              *
'*   Version 1.01.004 (September 10, 2003)                                      *
'*	 - Enhancement - added standardized remote debugging capability				*
'*   - Image change attribute									                *
'*   - Updated listing to show common attribute choices in Product Manager v2   *
'*                                                                              *
'*   Version 1.1.03                                                             *
'*   - Added showing price to single select for search_results                  *
'*                                                                              *
'*   Version 1.1.02                                                             *
'*   - Updated to work with MPO						                            *
'*                                                                              *
'*   Version 1.1                                                                *
'*   - Updated to work with Pricing Level Manager                               *
'*																				*
'*	This file should be located in the sfLib folder								*
'*	incGeneral.asp should include this file										*
'*																				*
'*   The contents of this file are protected by United States copyright laws	*
'*   as an unpublished work. No part of this file may be used or disclosed		*
'*   without the express written permission of Sandshot Software.				*
'*																				*
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.				*
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	Const cstrssTextBasedAttributeEmpty = "None"			'text to be used in the event a text based attribute is empty
	Const cblnssDisplayFullPriceInSingleAttributes = False	'set to true to display full price in attribute drop-down; applies only to items with a single attribute category

	Const cstrBaseImagePath = "images/"
	Const cstrImageSuffix = ".jpg"

	'Multiple Product Ordering CONFIGURATION
	cstrSSMPOAttributeDelimiter = "IMPOI"

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'For legacy reasons attribute IDs are built as strAttributeText & cstrSSTextBasedAttributeDelimiter & lngAttributeID


'**********************************************************
'*	Page Level variables
'**********************************************************

Const cstrSSTextBasedAttributeDelimiter = "|AttrExt|"		'Separator used between attribute and attribute ID in temporary table
Const cstrSSTextBasedAttributeHTMLDelimiter = "IAttrExtI"	'Separator used between attribute number (position) and attribute ID on search_results and detail pages. For text based attributes
Const cstrMultipleAttributeDelimiter = "/\"					'used in incAE GetAttName function to protect against commas in text based attributes
Const cblnUsessPricingLevels = False

'Enumerations
'These must match the order set in sfProductAdmin_custom
Const enAttrDisplay_Select = 0
Const enAttrDisplay_Radio = 1
Const enAttrDisplay_Text = 2
Const enAttrDisplay_TextOpt = 3
Const enAttrDisplay_Textarea = 4
Const enAttrDisplay_TextareaOpt = 5
Const enAttrDisplay_Checkbox = 6
Const enAttrDisplay_SelectShowPrice = 7
Const enAttrDisplay_SelectChangeImage = 8
Const enAttrDisplay_SelectChangePrice = 9
Const enAttrDisplay_RadioChangePrice = 10
Const enAttrDisplay_AttrAsQtyBox = 11
Const enAttrDisplay_RadioAttributePrice = 12
Const enAttrDisplay_Custom = 13
Const enAttrDisplay_Custom1 = 14

Dim cstrSSMPOAttributeDelimiter
Dim mdblConfiguredPrice
Dim mstrssAttributeExtenderjsOut
Dim mblnFormContentsWritten:	mblnFormContentsWritten = False

'**********************************************************
'*	Functions
'**********************************************************


'**********************************************************
'*	Begin Page Code
'**********************************************************

'**********************************************************
'*	Begin Function Definitions
'**********************************************************

Function emptyCustomerDefinedOptionText()
	emptyCustomerDefinedOptionText = ""
End Function	'emptyCustomerDefinedOptionText

'**********************************************************************************************************

Function isCustomerDefinedOption(byVal attrDisplayType)
	Select Case attrDisplayType
		Case enAttrDisplay_Text, enAttrDisplay_TextOpt, enAttrDisplay_Textarea, enAttrDisplay_TextareaOpt:
			isCustomerDefinedOption = True
		Case Else
			isCustomerDefinedOption = False
	End Select
End Function	'isCustomerDefinedOption

'**********************************************************************************************************

Function BuildAttribute(ByVal lngAttributeID, byVal strAttributeText)

  	If Len(Trim(strAttributeText & "")) = 0 Then
  		BuildAttribute = lngAttributeID
  	Else
  		BuildAttribute = strAttributeText & cstrSSTextBasedAttributeDelimiter & lngAttributeID
  	End If

End Function	'BuildAttribute

'**********************************************************************************************************

Function CleanAttributeArray(ByVal aryAttributes)

Dim i
Dim paryAttIDs()
Dim paryTemp
Dim plngCount
Dim pstrTempID

	If isArray(aryAttributes) Then
		If vDebug = 1 Then Response.Write "<fieldset><legend>CleanAttributeArray</legend>"
		plngCount = UBound(aryAttributes)
		ReDim paryAttIDs(plngCount)
		For i = 0 To plngCount
			If vDebug = 1 Then Response.Write "Attribute " & i & " of (" & plngCount & ") - " & aryAttributes(i) & "<BR>"
			If Len(aryAttributes(i)) > 0 Then
				paryTemp = Split(aryAttributes(i), cstrSSTextBasedAttributeDelimiter)
				If UBound(paryTemp) = 1 Then
					paryAttIDs(i) = paryTemp(1)
				Else
					paryAttIDs(i) = paryTemp(0)
				End If
			Else
				paryAttIDs(i) = ""
			End If
		Next 'i
		If vDebug = 1 Then Response.Write "</fieldset>"
	End If	'isArray(aryAttributes)

	CleanAttributeArray = paryAttIDs

End Function	'CleanAttributeArray

'**********************************************************************************************************

Sub CleanAttributeID(ByRef strID)
'Takes in string of attribute IDs which may be corrupted by Attribute Extender
'Returns comma delimited string of attribute IDs - Note because string is comma-delimited one has to protect against situation where commas appear

Dim paryAttributes
Dim paryTemp
Dim i
Dim pstrTempID
Dim pstrOut

	If vDebug = 1 Then
		Response.Write "CleanAttributeID - (strID) " & strID & "<BR>"
		Response.Write "cstrMultipleAttributeDelimiter " & cstrMultipleAttributeDelimiter & "<BR>"
	End If
	
	paryAttributes = Split(strID, cstrMultipleAttributeDelimiter)	'split on the attribute delimiter
	For i = 0 To UBound(paryAttributes)
		If vDebug = 1 Then Response.Write "Attribute " & i & " of (" & UBound(paryAttributes) & ") - " & paryAttributes(i) & "<BR>"
		paryTemp = Split(paryAttributes(i), cstrSSTextBasedAttributeDelimiter)	'split on comma because that is what SF does by default
		If UBound(paryTemp) = 1 Then
			pstrTempID = paryTemp(1)
		Else
			pstrTempID = paryTemp(0)
		End If

		If Len(pstrOut) = 0 Then
			pstrOut = pstrTempID
		Else
			pstrOut = pstrOut & "," & pstrTempID
		End If
	Next 'i
	
	If vDebug = 1 Then Response.Write "CleanAttributeID - (pstrOut) " & pstrOut & "<BR>"

	strID = pstrOut

End Sub	'CleanAttributeID

'**********************************************************************************************************

Function GetAttributeID(ByVal strAttr)

Dim plngPos

	plngPos = Instr(1, strAttr, cstrSSTextBasedAttributeDelimiter)
	If plngPos > 0 Then
		GetAttributeID = Right(strAttr,len(strAttr) - plngPos - Len(cstrSSTextBasedAttributeDelimiter) + 1)
	Else
		GetAttributeID = strAttr
	End If
			
End Function	'GetAttributeID

'**********************************************************************************************************

Function GetAttributeValue(ByVal strAttr)

Dim plngPos

	plngPos = Instr(1, strAttr, cstrSSTextBasedAttributeDelimiter)
	If plngPos > 0 Then
		GetAttributeValue = Left(strAttr, plngPos - 1)
	Else
		GetAttributeValue = ""
	End If
			
End Function	'GetAttributeValue

'**********************************************************************************************************

Function isAttributeMatch(ByVal strAttrTest, ByVal lngAttrID, ByVal strAttrValue)

Dim plngPos

	plngPos = Instr(1, strAttrTest, cstrSSTextBasedAttributeDelimiter)
	If plngPos > 0 Then
		isAttributeMatch = CBool(strAttrTest = (strAttrValue & cstrSSTextBasedAttributeDelimiter & lngAttrID))
	Else
		isAttributeMatch = CBool(CStr(strAttrTest & "") = CStr(lngAttrID & ""))
	End If
	
	If vDebug = 1 Then
		Response.Write "<fieldset><legend>isAttributeMatch</legend>"
		Response.Write "strAttrTest: " & strAttrTest & "<br>"
		Response.Write "lngAttrID: " & lngAttrID & "<br>"
		Response.Write "strAttrValue: " & strAttrValue & "<br>"
		Response.Write "isAttributeMatch: " & isAttributeMatch & "<br>"
		Response.Write "</fieldset>"
	End If
			
End Function	'isAttributeMatch

'**********************************************************************************************************

Sub GetAttributes(byVal strProductID, byRef aProdAttr, byRef iProdAttrNum)
'Purpose: fills aProdAttr with attribute selection from Request.Form
'Returns: Nothing
'		  sets values for aProdAttr which is byRef

ReDim aProdAttr(iProdAttrNum)
Dim i
Dim pstrFieldName
Dim pstrNewVar		
Dim pstrTempAttribute
Dim vItem	

	If vDebug = 1 Then
		Response.Write "<h4>mblnMPOPage_ssMPO: " & mblnMPOPage_ssMPO & "</h4>"
		Response.Write "<h4>iProdAttrNum: " & iProdAttrNum & "</h4>"
	End If
		
	If mblnMPOPage_ssMPO Then
		For i = 1 to iProdAttrNum
			
			pstrNewVar = "attr" & i & cstrSSMPOAttributeDelimiter & strProductID		    		     
			pstrTempAttribute = Request.Form(pstrNewVar)
			If vDebug = 1 Then
				Response.Write "Looking for attribute " & i & " of " & iProdAttrNum & "<br>"
				Response.Write pstrNewVar & ": " & pstrTempAttribute & "<br>"
			End If
			If (Len(pstrTempAttribute) > 0) Then	
				aProdAttr(i-1) = pstrTempAttribute
			Else
				'check to see if attribute extender
				
				pstrFieldName = pstrNewVar & cstrSSTextBasedAttributeHTMLDelimiter
				If vDebug = 1 Then Response.Write "<b>Looking for </b>: <i>" & pstrFieldName & "</i><br>"
				For Each vItem in Request.Form
					If instr(1, vItem, pstrFieldName) > 0 Then
						pstrTempAttribute = Request.Form(vItem)  & Replace(vItem, pstrNewVar & cstrSSTextBasedAttributeHTMLDelimiter, cstrSSTextBasedAttributeDelimiter)
						aProdAttr(i-1) = pstrTempAttribute
						If vDebug = 1 Then Response.Write "pstrNewVar - " & pstrNewVar & ": " & pstrTempAttribute & "<br>"
						Exit For
					End If
				Next			
			End If							
		Next	 		         
	Else
		For i = 1 to iProdAttrNum
			pstrNewVar = "attr" & i				    		     
			pstrTempAttribute = Request.Form(pstrNewVar)
			If (pstrTempAttribute <> "") Then	
				aProdAttr(i-1) = pstrTempAttribute
			Else
				For Each vItem in Request.Form
					If instr(1,vItem,pstrNewVar & cstrSSTextBasedAttributeHTMLDelimiter) > 0 Then
						pstrTempAttribute = Request.Form(vItem) & Replace(vItem, pstrNewVar & cstrSSTextBasedAttributeHTMLDelimiter, cstrSSTextBasedAttributeDelimiter)
						aProdAttr(i-1) = pstrTempAttribute
						If vDebug = 1 Then Response.Write "" & pstrNewVar & ": " & pstrTempAttribute & "<br>"
						Exit For
					End If
				Next			
			End If							
		Next	 	
	End If	'mblnMPOPage_ssMPO
	
	If vDebug = 1 Then 
		Response.Write "<fieldset><legend>Attributes for product </b>: <i>" & strProductID & "</i></legend>"
		For i = 1 to iProdAttrNum
			Response.Write "&nbsp;&nbsp;<i>Attribute " & i & ": " & aProdAttr(i-1) & "</i><br>"
		Next
		Response.Write "</fieldset>" 	
	End If

End Sub	'GetAttributes

'**********************************************************************************************************

Function getAttrDetails(iAttrID)

Dim sLocalSQL, rsFindAttr, aLocalAttr
Dim plngPos,plngNewAttrID

	plngPos = Instr(1,iAttrID,cstrSSTextBasedAttributeDelimiter)
	If plngPos > 0 Then
		plngNewAttrID = Right(iAttrID,len(iAttrID) - plngPos - Len(cstrSSTextBasedAttributeDelimiter) + 1)
		If cblnUsessPricingLevels Then
		'	sLocalSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrdtID = " & iAttrID
			sLocalSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType, attrdtPLPrice FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrdtID = " & plngNewAttrID
		Else
			sLocalSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrdtID = " & plngNewAttrID
		End If
	Else
		If cblnUsessPricingLevels Then
			sLocalSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType, attrdtPLPrice FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrdtID = " & iAttrID
		Else
			sLocalSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrdtID = " & iAttrID
		End If
	End If
	
	Set rsFindAttr = CreateObject("ADODB.RecordSet")
	rsFindAttr.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
	If vDebug = 1 Then Response.Write "getAttrDetails - rsFindAttr.RecordCount: " & rsFindAttr.RecordCount & "<BR>"
	If Not rsFindAttr.EOF Then
		Redim aLocalAttr(4)
		If plngPos > 0 Then
			If Len(Left(iAttrID,plngPos-1)) = 0 Then
				aLocalAttr(0) = rsFindAttr.Fields("attrdtName") & " " & cstrssTextBasedAttributeEmpty
			Else
				aLocalAttr(0) = rsFindAttr.Fields("attrdtName") & " " & Left(iAttrID,plngPos-1)
			End If
		Else
			aLocalAttr(0) = rsFindAttr.Fields("attrdtName")
			aLocalAttr(0) = rsFindAttr.Fields("attrName") & ": " & rsFindAttr.Fields("attrdtName")
		End If
		If cblnUsessPricingLevels Then
			aLocalAttr(1) = GetPricingLevelPrice(rsFindAttr.Fields("attrdtPrice"), rsFindAttr.Fields("attrdtPLPrice"))
			If cblnDebugPricingLevel Then debugprint "attrdtPrice",rsFindAttr.Fields("attrdtPrice")
			If cblnDebugPricingLevel Then debugprint "attrdtPLPrice",rsFindAttr.Fields("attrdtPLPrice")
		Else
			aLocalAttr(1) = rsFindAttr.Fields("attrdtPrice")
		End If

		aLocalAttr(2) = rsFindAttr.Fields("attrdtType")
		aLocalAttr(3) = rsFindAttr.Fields("attrName")
		
	End If
		
	closeObj(rsFindAttr)
	getAttrDetails = aLocalAttr
	
End Function	'getAttrDetails

'**********************************************************************************************************

Function HasAttributes()

Dim pblnHasAttributes
Dim vItem

	'On Error Resume Next

	pblnHasAttributes = (Trim(Request.Form("attr1")) = "")
	If Not pblnHasAttributes Then
		For Each vItem in Request.Form
			If Len(Trim(Request.Form(vItem))) > 0 Then
				pblnHasAttributes = True
				Exit For
			End If
		Next
	End If
	
	HasAttributes = pblnHasAttributes

End Function	'HasAttributes

'**********************************************************************************************************

Function AttributeImageName(ByRef objRS, ByVal strFormName)

Dim pstrTemp

	pstrTemp = cstrBaseImagePath & Trim(objRS.Fields("attrdtId").Value & "") + cstrImageSuffix
	If Len(cstrattrdtImage_Field) > 0 Then
		If Len(objRS.Fields(cstrattrdtImage_Field).Value & "") > 0 Then
			pstrTemp = objRS.Fields(cstrattrdtImage_Field).Value
		End If
	End If
	
	AttributeImageName = pstrTemp
					
End Function	'AttributeImageName

'**********************************************************************************************************

Function attributeLink_new(ByRef aryAttribute, ByVal strName, ByVal blnAttribute)

Dim pstrOut
Dim pstrURL
Dim pstrExtra

	If blnAttribute Then
		pstrExtra = aryAttribute(enAttribute_Extra)
		If Len(pstrExtra) > 0 Then
			pstrOut = pstrExtra
		Else
			pstrURL = aryAttribute(enAttribute_URL)
			If Len(pstrURL) > 0 Then
				pstrOut = pstrOut & "<a class=attExtAttributeAnchor href=" & pstrURL & " " & pstrExtra & ">" & strName & "</a>"
			Else
				pstrOut = pstrOut & "<span class=""attributeCategoryName"">" & strName & "</span>"
			End If
		End If
	Else
		pstrExtra = aryAttribute(enAttributeDetail_Extra)
		If Len(pstrExtra) > 0 Then
			pstrOut = pstrExtra
		Else
			pstrURL = aryAttribute(enAttributeDetail_URL)
			If Len(pstrURL) > 0 Then
				pstrOut = pstrOut & "<a class=attExtAttributeDetailAnchor href=" & pstrURL & " " & pstrExtra & ">" & strName & "</a>"
			Else
				pstrOut = pstrOut & "<span class=""attributeCategoryName"">" & strName & "</span>"
			End If
		End If
	End If
	
	attributeLink_new = pstrOut

End Function	'attributeLink_new

'**********************************************************************************************************

Function setImageOnLoad_Detail(byRef objRSAttributes, byVal strProdID)

Dim i
Dim paryAttributes
Dim pblnImageChangeScript
Dim pbytattrDisplayStyle

	pblnImageChangeScript = False
	
	If isObject(objRSAttributes) Then
		With objRSAttributes
			If .State <> 1 Then Exit Function
			If Not .EOF Then
				Do While Not .EOF
					pbytattrDisplayStyle = .Fields("attrDisplayStyle").Value
					If Len(pbytattrDisplayStyle & "") = 0 Then pbytattrDisplayStyle = enAttrDisplay_Select
					
					If pbytattrDisplayStyle = enAttrDisplay_SelectChangeImage Then
						pblnImageChangeScript = True
						Exit Do
					End If
					.MoveNext
				Loop
				.MoveFirst
			End If	'objRSAttributes.EOF
		End With
	ElseIf getProductInfo(txtProdId, enProduct_AttrNum) > 0 Then
		paryAttributes = getProductInfo(txtProdId, enProduct_attributes)
		If isArray(paryAttributes) Then
			For i = 0 To UBound(paryAttributes)
				pbytattrDisplayStyle = CLng(paryAttributes(i)(enAttribute_DisplayStyle))
				
				If pbytattrDisplayStyle = enAttrDisplay_SelectChangeImage Then
					pblnImageChangeScript = True
					Exit For
				End If
			Next 'i
		End If
		response.Write "<h1>Test</h1>"
	End If
	
	If pblnImageChangeScript Then setImageOnLoad_Detail = " onload=" & Chr(34) & "setCustomImage(this, '" & strProdID & "');" & Chr(34)

End Function	'setImageOnLoad_Detail

'**********************************************************************************************************

Function setImageOnLoad_SearchResults(byRef arrAtt, byVal strProdID)

Dim i
Dim pblnImageChangeScript
Dim pbytattrDisplayStyle

	pblnImageChangeScript = False
	
	If isArray(arrAtt) Then
		For i = 0 To UBound(arrAtt, 2)
			pbytattrDisplayStyle = arrAtt(3, i)
			If Len(pbytattrDisplayStyle & "") = 0 Then pbytattrDisplayStyle = enAttrDisplay_Select
		
			If pbytattrDisplayStyle = enAttrDisplay_SelectChangeImage Then
				pblnImageChangeScript = True
				Exit For
			End If
		Next 'i
	End If	'isArray(arrAtt)
	
	If pblnImageChangeScript Then setImageOnLoad_SearchResults = " onload=" & Chr(34) & "setCustomImage(this, '" & strProdID & "');" & Chr(34)

End Function	'setImageOnLoad_SearchResults

'**********************************************************************************************************

Sub DisplayAttributes(ByRef strOut, ByVal strFormName)

Dim pbytattrDisplayStyle
Dim pstrattrName		
Dim pblnSelectOpen
Dim pblnTableOpen
Dim pblnChecked
Dim pblnDefault
Dim pstrAttrTitle
Dim pstrFieldName
Dim pstrSelectedStyle
Dim pstrOnChangeText
Dim pstrOptionTemplate
Dim pstrOptionLineOut

Dim pstrAttrDisplay
Dim pstrAttrDisplayAlt

Dim pdblBasePrice
Dim pdblAttrDelta
Dim pstrAttrPricing
Dim plngAttrPosition

	pblnSelectOpen = False
	pblnTableOpen = False
	pblnChecked = False
	
	pdblBasePrice = getProductInfo(txtProdId, enProduct_SellPrice)
	
	mdblConfiguredPrice = pdblBasePrice
	pstrAttrPricing = ""

	Do While Not rsProdAttributes.EOF
		attrName = rsProdAttributes.Fields("attrName").Value
		pbytattrDisplayStyle = rsProdAttributes.Fields("attrDisplayStyle").Value
		If Len(pbytattrDisplayStyle & "") = 0 Then pbytattrDisplayStyle = enAttrDisplay_Select
		strAttrPrice = ""
		
		Select Case rsProdAttributes.Fields("attrdtType").Value
			Case 1
				pdblAttrDelta = rsProdAttributes.Fields("attrdtPrice").Value
			Case 2
				pdblAttrDelta = -1 * rsProdAttributes.Fields("attrdtPrice").Value
			Case Else
				pdblAttrDelta = 0
		End Select

		If Len(cstrattrdtDefault_Field) > 0 Then
			If Len(rsProdAttributes.Fields(cstrattrdtDefault_Field).Value & "") = 0 Then
				pblnDefault = CBool(Trim(attrName) <> Trim(attrNamePrev))
			Else
				pblnDefault = CBool(rsProdAttributes.Fields(cstrattrdtDefault_Field).Value)
			End If
		Else
			pblnDefault = CBool(Trim(attrName) <> Trim(attrNamePrev))
		End If

		If pblnDefault Then
			pstrSelectedStyle = " style='background-color:red;'"
			pstrSelectedStyle = " style='background-color:lightgrey;border-left-width:1px; border-right-width:1px; border-top-style:dotted; border-top-width:1px; border-bottom-style:dotted; border-bottom-width:1px'"
			pstrSelectedStyle = " class='attExtDivDefault'"
			pblnChecked = False
			mdblConfiguredPrice = CDbl(mdblConfiguredPrice) + CDbl(pdblAttrDelta)
		Else
			pstrSelectedStyle = " style='background-color:blue;'"
			pstrSelectedStyle = " style='background-color:white;border-left-width:1px; border-right-width:1px; border-top-style:dotted; border-top-width:1px; border-bottom-style:dotted; border-bottom-width:1px'"
			pstrSelectedStyle = " class='attExtDiv'"
		End If

'available items:
'Attribute
'- Attribute Name
'- URL
'- Extra
'- Display Style

'Attribute detail
'- Name
'- URL
'- Extra
'- Image Path
'- Selected
'- Price
'- PriceType
'- Weight

'Start new attribute
'repeating item - odd
'repeating item - even; defaults to odd if not present
'repeating item - selected; defaults to odd/even if not present
'Close attribute

'Example
'Opening
'<tr><td align='left'>{attributeDisplayText}</td>
'<td><select style="{attributeStyle}" name='{attributeName}' id='{attributeName}'>
'Repeating
'<option style="{attributeDetailStyle}" value="{attributeDetailValue}">{attributeDetailDisplayText}<option>
'Closing
'</select></td></tr>

'Example
'Opening
'<tr><td align='left'>
'<table border=''>
'<tr><td colspan=2></td></tr>
'<tr><td>{attributeImage}</td><td>
'<table border=''>
'Repeating
'<tr><td>{attribute}</td></tr>
'Closing
'</table>
'</td></tr>
'</table>
'</td></tr>
'<td><select style="{attributeStyle}" name='{attributeName}' id='{attributeName}'>

'Repeating
'<option style="{attributeDetailStyle}" value="{attributeDetailValue}">{attributeDetailDisplayText}<option>
'</select></td>


		If Trim(attrName) <> Trim(attrNamePrev) Then
			plngAttrPosition = 0
			If iCounter > 0 Then
				If pblnSelectOpen Then
					strOut = strOut & "</select>" & vbcrlf
					pblnChecked = False
					pblnSelectOpen = False
				End If
				strOut = strOut & vbcrlf _
								& "<script language='javascript'  type='text/javascript'>" & vbcrlf _
								& "	prodBasePrice = " & pdblBasePrice & ";" & vbcrlf _
								& "	cstrBaseImagePath = " & Chr(34) & getProductInfo(txtProdId, enProduct_ImageSmallPath) & txtProdId & Chr(34) & ";" & vbcrlf _
								& pstrAttrPricing _
								& "</script>"
				strOut = strOut & "</td></tr>"
				pstrAttrPricing = ""
			End If	'iCounter > 0
			pstrattrName = "attr" & (iAttrCounter(iAttrNum))
			
			'Write Attribute Category Name
			Select Case pbytattrDisplayStyle
				Case 998:	'Dummy for Attribute Category in-line
					strOut = strOut & "<tr><td align=right valign=top>" _
						& attributeLink(rsProdAttributes, attrName, True) _
						& "</td><td>" & vbcrlf
				Case 999:	'Dummy for Attribute Category block
					strOut = strOut & "<tr><td align=left valign=top>" _
						& attributeLink(rsProdAttributes, attrName, True) _
						& "</td></tr><tr><td align=left valign=top>" & vbcrlf
				Case enAttrDisplay_RadioAttributePrice:
					'Do Not Display Attribute Category Name
					strOut = strOut & "<tr><td>" & vbcrlf
				Case Else:
					'default to attribute category on separate line
					strOut = strOut & "<tr><td align=left valign=top>" _
						& attributeLink(rsProdAttributes, attrName, True) _
						& "</td></tr><tr><td align=left valign=top>" & vbcrlf
			End Select


			Select Case pbytattrDisplayStyle
				Case enAttrDisplay_Select, enAttrDisplay_SelectShowPrice
					strOut = strOut & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "'>" & vbcrlf
					pblnSelectOpen = True
				Case enAttrDisplay_SelectChangeImage
					pstrOnChangeText = "onchange=""changeCustomImage('" & EscapeFormName(strFormName) & "', this);"""
					strOut = strOut & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "' onchange=""changeCustomImage('" & EscapeFormName(strFormName) & "', this);"">" & vbcrlf
					pblnSelectOpen = True
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "maryAttrImages['" & EscapeFormName(strFormName) & "'] = new Array();" & vbcrlf
				Case enAttrDisplay_SelectChangePrice
					pstrOnChangeText = "onchange='updateProductPrice(this," & iAttrNum & ");'"
					strOut = strOut & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "' " & pstrOnChangeText & ">" & vbcrlf
					pblnSelectOpen = True
					'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "maryAttrImages['" & EscapeFormName(strFormName) & "'] = new Array();" & vbcrlf
				Case enAttrDisplay_AttrAsQtyBox
					strOut = strOut & "<input type=hidden name=ssAttrQTY id=ssAttrQTY value=" & iAttrNum & ">" _
							& "<table cellpadding=2 cellspacing=0 border=1>"
					pblnTableOpen = True
				Case enAttrDisplay_RadioAttributePrice
					strOut = strOut & "<table class=""attributeDisplay"">"
					pblnTableOpen = True
				Case enAttrDisplay_Radio, enAttrDisplay_RadioChangePrice:	'radio
					pblnChecked = False	'this is necessary so first radio option is selected by default if multiple radio options are used
			End Select
			pstrOnChangeText = ""
		End If	'Trim(attrName) <> Trim(attrNamePrev)

		Select Case rsProdAttributes.Fields("attrdtType").Value
			Case 1
				pstrAttrDisplayAlt = " " & FormatCurrency( CDbl(pdblBasePrice) + CDbl(rsProdAttributes.Fields("attrdtPrice").Value)) & ""
				pstrAttrDisplay = " (Add " & FormatCurrency(rsProdAttributes("attrdtPrice")) & ")"
				pdblAttrDelta = rsProdAttributes.Fields("attrdtPrice").Value
			Case 2
				pstrAttrDisplayAlt = " " & FormatCurrency( CDbl(pdblBasePrice) - CDbl(rsProdAttributes.Fields("attrdtPrice").Value)) & ""
				pstrAttrDisplay = " (Subtract " & FormatCurrency(rsProdAttributes("attrdtPrice")) & ")"
				pdblAttrDelta = -1 * rsProdAttributes.Fields("attrdtPrice").Value
			Case Else
				pstrAttrDisplay = ""
				pstrAttrDisplayAlt = " " & FormatCurrency(pdblBasePrice) & ""
				pdblAttrDelta = 0
		End Select

		plngAttrPosition = plngAttrPosition + 1
		If Len(pstrAttrPricing) = 0 Then
			pstrAttrPricing = "setCurrentAttributePrice(" & iAttrNum & "," & pdblAttrDelta & ");" & vbcrlf
		End If
		pstrAttrPricing = pstrAttrPricing & "setAttributePrice(" & iAttrNum & "," & plngAttrPosition & "," & pdblAttrDelta & ");" & vbcrlf

		If getProductInfo(txtProdId, enProduct_AttrNum) = 1 And cblnssDisplayFullPriceInSingleAttributes Then
			strAttrPrice = pstrAttrDisplayAlt
		Else
			strAttrPrice = pstrAttrDisplay
		End If
		
		pstrAttrTitle = Trim(rsProdAttributes.Fields("attrdtName").Value & "")
		'pstrAttrTitle = ""
		
		pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & rsProdAttributes.Fields("attrdtId").Value
		
		Select Case pbytattrDisplayStyle
			'0;Combo;1;Radio;2;Text - required;3;Text - optional;4;Textarea - required;5;Textarea - optional;6;Checkbox
			Case enAttrDisplay_Select: 'select
					If pblnChecked Then
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
					Else
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """ selected>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
						pblnChecked = True
					End If
			Case enAttrDisplay_Radio, enAttrDisplay_RadioChangePrice:	'radio
					If pbytattrDisplayStyle = enAttrDisplay_RadioChangePrice Then
						pstrOnChangeText = " onclick='updateProductPrice(this," & iAttrNum & "," & plngAttrPosition & ");'"
					Else
						pstrOnChangeText = ""
					End If
					
					If pblnChecked Then
						strOut = strOut & "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
					Else
						strOut = strOut & "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & rsProdAttributes.Fields("attrdtId").Value & """ checked>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
						pblnChecked = True
					End If
			Case enAttrDisplay_RadioAttributePrice:	'radio
					pstrOptionTemplate = "<div {SelectedStyle}>" _
									   & "<input type=""radio"" name=""{radioName}"" id=""{radioID}"" {OnChangeText} value=""{attrdtId}""{checked}>" _
									   & "<label for=""{radioID}"">" _
									   & "<div class=""attributeDisplay"">" _
									   & "{AttrTitle}<br />" _
									   & "Rec. Retail: {attrdtExtra}<br />" _
									   & "<span class=""SalesPrice"">Our Price: {AttrPrice}</span>" _
									   & "</div>" _
									   & "</label>" _
									   & "</div>"

					pstrOptionTemplate = "<tr>" _
									   & "<td {SelectedStyle} valign=""top""><input type=""radio"" name=""{radioName}"" id=""{radioID}"" {OnChangeText} value=""{attrdtId}""{checked}></td>" _
									   & "<td {SelectedStyle} nowrap><label for=""{radioID}"">{AttrTitle}<br />" _
									   & "Rec. Retail: {attrdtExtra}&nbsp;&nbsp;<br />" _
									   & "<span class=""SalesPrice"">Our Price: {AttrPrice}</span>" _
									   & "</label>" _
									   & "</td>" _
									   & "</tr>"

					If pblnDefault Then
						pstrOptionLineOut = Replace(pstrOptionTemplate, "{SelectedStyle}", " class=""attributeDisplaySelected""")
					Else
						pstrOptionLineOut = Replace(pstrOptionTemplate, "{SelectedStyle}", " class=""attributeDisplay""")
					End If
					
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{radioName}", pstrattrName)
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{radioID}", pstrattrName & rsProdAttributes.Fields("attrdtId").Value)
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{OnChangeText}", pstrOnChangeText)
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{attrdtId}", rsProdAttributes.Fields("attrdtId").Value)
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{AttrTitle}", pstrAttrTitle)
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{attributeLink}", attributeLink(rsProdAttributes, pstrAttrTitle, False))
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{AttrPrice}", FormatCurrency(pdblAttrDelta, 2))
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{attrdtExtra}", FormatCurrency(rsProdAttributes.Fields("attrdtExtra").Value, 2))
					pstrOptionLineOut = Replace(pstrOptionLineOut, "{Counter}", iCounter)

					If iCounter > 0 Then
						'pstrOptionLineOut = "<tr><td class=""attributeDisplay"" colspan=""2""><hr /></td></tr>" & pstrOptionLineOut
						pstrOptionLineOut = "<tr class=""attributeDisplay""><td></td><td><hr /></td></tr>" & pstrOptionLineOut
					End If

					If pblnChecked Then
						'strOut = strOut & "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
						strOut = strOut & Replace(pstrOptionLineOut, "{checked}", "")
					Else
						'strOut = strOut & "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & rsProdAttributes.Fields("attrdtId").Value & """ checked>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
						strOut = strOut & Replace(pstrOptionLineOut, "{checked}", " checked")
						pblnChecked = True
					End If
			Case enAttrDisplay_Text:	'text - required
					strOut = strOut & "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
			Case enAttrDisplay_TextOpt:	'text - optional
					strOut = strOut & "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
					'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeOptional(""" & strFormName & """, """ & pstrFieldName & """, """ & Server.HTMLEncode(pstrAttrTitle) & """);" & vbcrlf

					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = true;" & vbcrlf
			Case enAttrDisplay_Textarea:	'textarea - required
					strOut = strOut & "<textarea rows='2' columns='40' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''></textarea>" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
			Case enAttrDisplay_TextareaOpt:	'textarea - optional
					strOut = strOut & "<textarea rows='2' columns='40' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''></textarea>" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = 'true';" & vbcrlf
			Case enAttrDisplay_Checkbox:	'checkbox
					strOut = strOut & "<input type='checkbox' name='" & pstrFieldName & "' value='on'>" & pstrAttrTitle & strAttrPrice & "<br>" & vbcrlf
			Case enAttrDisplay_SelectShowPrice:	'select (show price)
					If pblnChecked Then
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & pstrAttrTitle & pstrAttrDisplayAlt & "</option>" & vbcrlf
					Else
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """ selected>" & pstrAttrTitle & pstrAttrDisplayAlt & "</option>" & vbcrlf
						pblnChecked = True
					End If
			Case enAttrDisplay_SelectChangeImage:	'custom example
					If pblnChecked Then
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
					Else
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """ selected>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
						pblnChecked = True
					End If
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "maryAttrImages['" & EscapeFormName(strFormName) & "'][" & rsProdAttributes.Fields("attrdtId").Value & "] = " & Chr(34) & AttributeImageName(rsProdAttributes, EscapeFormName(strFormName)) & Chr(34) & ";" & vbcrlf
			Case enAttrDisplay_AttrAsQtyBox:	'attribute as quanity box - assumes MPO installed
					strOut = strOut & "<tr><td>" & pstrAttrTitle & "</td>" _
									& "<td>" _
									& " <input type=hidden name='" & pstrattrName & "' value='" & rsProdAttributes.Fields("attrdtId").Value & "'>" _
									& " <input type=text name='QUANTITY' size=4 onblur='return validQuantity(this);'>" _
									& "</td></tr>" & vbcrlf
			Case enAttrDisplay_Custom:	'custom example
					strOut = strOut & "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
                    strOut = strOut & "<a href=""#"" onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('button_r1_c1','','images/button_r1_c1_f2.gif',1);"" ><img name=""button_r1_c1"" src=""images/button_r1_c1.gif"" width=""187"" height=""43"" border=""0"" onMouseDown=""MM_openBrWindow('/charm_directorynewmb.asp','','width=550,height=500')""></a></td>"
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = 'true';" & vbcrlf
					mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "mtxtCharm = document." & strFormName & "." & pstrFieldName & ";" & vbcrlf
			Case Else: 
					If pblnChecked Then
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
					Else
						strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """ selected>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
						pblnChecked = True
					End If
					'strOut = strOut & "<option value=""" & rsProdAttributes.Fields("attrdtId").Value & """>" & pstrAttrTitle & strAttrPrice & "</option>"
		End Select

		attrNamePrev = attrName
		icounter = icounter + 1
		rsProdAttributes.MoveNext
	Loop
				
	If pblnSelectOpen Then
		strOut = strOut & "</select>" & vbcrlf
		pblnSelectOpen = False
	End If
	
	If pblnTableOpen Then
		strOut = strOut & "</table>" & vbcrlf
		pblnTableOpen = False
	End If
	
	If Len(pstrAttrPricing) > 0 Then
	strOut = strOut & vbcrlf _
					& "<tr><td><script language='javascript'  type='text/javascript'>" & vbcrlf _
					& pstrAttrPricing _
					& "</script></td></tr>"
	End If

End Sub	'DisplayAttributes

'**********************************************************************************************************

Sub DisplayAttributes_New(ByRef strOut, ByVal strFormName)

Dim i, j
Dim paryAttributes
Dim pblnChecked
Dim pblnDefault
Dim pbytattrDisplayStyle
Dim pdblBasePrice
Dim pstrattrName		
Dim pstrAttrPricing
Dim pstrClosingTag
Dim pobjStringBuilder


Dim pstrAttrTitle
Dim pstrFieldName
Dim pstrSelectedStyle
Dim pstrOnChangeText
Dim pstrOptionTemplate
Dim pstrOptionLineOut
Dim pstrAttrDisplay
Dim pstrAttrDisplayAlt
Dim pdblAttrDelta

	pblnChecked = False
	
	pdblBasePrice = getProductInfo(txtProdId, enProduct_SellPrice)
	mdblConfiguredPrice = pdblBasePrice
	pstrAttrPricing = ""
	
	Set pobjStringBuilder = New FastString
	pobjStringBuilder.Append vbcrlf _
							 & "<script language='javascript'  type='text/javascript'>" & vbcrlf _
							 & "	prodBasePrice = " & pdblBasePrice & ";" & vbcrlf _
							 & "	cstrBaseImagePath = " & Chr(34) & getProductInfo(txtProdId, enProduct_ImageSmallPath) & txtProdId & Chr(34) & ";" & vbcrlf _
							 & pstrAttrPricing _
							 & "</script>"

	paryAttributes = getProductInfo(txtProdId, enProduct_attributes)
	If isArray(paryAttributes) Then
	For i = 0 To UBound(paryAttributes)
		attrName = paryAttributes(i)(enAttribute_Name)
		pstrattrName = "attr" & CStr(i + 1)
		If Len(CStr(paryAttributes(i)(enAttribute_DisplayStyle))) = 0 Then paryAttributes(i)(enAttribute_DisplayStyle) = enAttrDisplay_Select
		pbytattrDisplayStyle = CLng(paryAttributes(i)(enAttribute_DisplayStyle))

		'Open the attribute display
		Select Case pbytattrDisplayStyle
			Case enAttrDisplay_Select, enAttrDisplay_SelectShowPrice
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" _
										 & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "'>" & vbcrlf
				pstrClosingTag = "</select>"
			Case enAttrDisplay_SelectChangeImage
				pstrOnChangeText = "onchange=""changeCustomImage('" & EscapeFormName(strFormName) & "', this);"""
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" _
										 & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "' onchange=""changeCustomImage('" & EscapeFormName(strFormName) & "', this);"">" & vbcrlf
				pstrClosingTag = "</select>"
				mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "maryAttrImages['" & EscapeFormName(strFormName) & "'] = new Array();" & vbcrlf
			Case enAttrDisplay_SelectChangePrice
				pstrOnChangeText = "onchange='updateProductPrice(this," & i + 1 & ");'"
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" _
										 & "<select style=""" & C_FORMDESIGN  & """ name='" & pstrattrName & "' " & pstrOnChangeText & ">" & vbcrlf
				pstrClosingTag = "</select>"
			Case enAttrDisplay_AttrAsQtyBox
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" _
										 & "<input type=hidden name=ssAttrQTY id=ssAttrQTY value=" & i + 1 & ">" _
										 & "<table cellpadding=2 cellspacing=0 border=1>"
				pstrClosingTag = "</table>"
			Case enAttrDisplay_RadioAttributePrice
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" _
										 & "<table class=""attributeDisplay"">"
				pstrClosingTag = "</table>"
			Case enAttrDisplay_RadioAttributePrice:
				'Do Not Display Attribute Category Name
				pobjStringBuilder.Append "<tr><td>" & vbcrlf
			Case enAttrDisplay_Radio, enAttrDisplay_RadioChangePrice:	'radio
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>"
				pblnChecked = False	'this is necessary so first radio option is selected by default if multiple radio options are used
			Case 998:	'Dummy for Attribute Category in-line
				pobjStringBuilder.Append "<tr><td align=right valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td><td>" & vbcrlf
			Case 999:	'Dummy for Attribute Category block
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" & vbcrlf
			Case Else:
				'default to attribute category on separate line
				pobjStringBuilder.Append "<tr><td align=left valign=top>" _
										 & attributeLink_new(paryAttributes(i), attrName, True) _
										 & "</td></tr><tr><td align=left valign=top>" & vbcrlf
		End Select

		'Now for the attribute details
		For j = 0 To UBound(paryAttributes(i)(enAttribute_DetailArray))
			Select Case paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Type)
				Case 1
					pdblAttrDelta = CDbl(paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Price))
					pstrAttrDisplayAlt = " " & FormatCurrency(CDbl(pdblBasePrice) + CDbl(pdblAttrDelta))
					pstrAttrDisplay = " (Add " & FormatCurrency(pdblAttrDelta) & ")"
				Case 2
					pdblAttrDelta = -1 * CDbl(paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Price))
					pstrAttrDisplayAlt = " " & FormatCurrency(CDbl(pdblBasePrice) + pdblAttrDelta)
					pstrAttrDisplay = " (Subtract " & FormatCurrency(paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Price)) & ")"
				Case Else
					pdblAttrDelta = 0
					pstrAttrDisplayAlt = " " & FormatCurrency(pdblBasePrice)
					pstrAttrDisplay = ""
			End Select
			
			'check to see if this is selected by default
			If Len(paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Default) & "") = 0 Then
				pblnDefault = CBool(Trim(attrName) <> Trim(attrNamePrev))
			Else
				pblnDefault = CBool(paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Default))
			End If
			If pblnDefault Then
				pstrSelectedStyle = " style='background-color:red;'"
				pstrSelectedStyle = " style='background-color:lightgrey;border-left-width:1px; border-right-width:1px; border-top-style:dotted; border-top-width:1px; border-bottom-style:dotted; border-bottom-width:1px'"
				pstrSelectedStyle = " class='attExtDivDefault'"
				mdblConfiguredPrice = CDbl(mdblConfiguredPrice) + CDbl(pdblAttrDelta)
				pblnChecked = False
			Else
				pstrSelectedStyle = " style='background-color:blue;'"
				pstrSelectedStyle = " style='background-color:white;border-left-width:1px; border-right-width:1px; border-top-style:dotted; border-top-width:1px; border-bottom-style:dotted; border-bottom-width:1px'"
				pstrSelectedStyle = " class='attExtDiv'"
			End If
			
			If Len(pstrAttrPricing) = 0 Then pstrAttrPricing = "setCurrentAttributePrice(" & i + i & "," & pdblAttrDelta & ");" & vbcrlf
			pstrAttrPricing = pstrAttrPricing & "setAttributePrice(" & i + i & "," & j + 1 & "," & pdblAttrDelta & ");" & vbcrlf

			If getProductInfo(txtProdId, enProduct_AttrNum) = 1 And cblnssDisplayFullPriceInSingleAttributes Then
				strAttrPrice = pstrAttrDisplayAlt
			Else
				strAttrPrice = pstrAttrDisplay
			End If
			
			pstrAttrTitle = paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Name)
			pstrFieldName = pstrattrName & cstrSSTextBasedAttributeHTMLDelimiter & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID)

			'Write out the actual attribute detail
			Select Case pbytattrDisplayStyle
				Case enAttrDisplay_Select: 'select
						If pblnChecked Then
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
						Else
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """ selected>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
							pblnChecked = True
						End If
				Case enAttrDisplay_Radio, enAttrDisplay_RadioChangePrice:	'radio
						If pbytattrDisplayStyle = enAttrDisplay_RadioChangePrice Then
							pstrOnChangeText = " onclick='updateProductPrice(this," & i + 1 & "," & j + 1 & ");'"
						Else
							pstrOnChangeText = ""
						End If
						
						If pblnChecked Then
							pobjStringBuilder.Append "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
						Else
							pobjStringBuilder.Append "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """ checked>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
							pblnChecked = True
						End If
				Case enAttrDisplay_RadioAttributePrice:	'radio
						pstrOptionTemplate = "<div {SelectedStyle}>" _
										& "<input type=""radio"" name=""{radioName}"" id=""{radioID}"" {OnChangeText} value=""{attrdtId}""{checked}>" _
										& "<label for=""{radioID}"">" _
										& "<div class=""attributeDisplay"">" _
										& "{AttrTitle}<br />" _
										& "Rec. Retail: {attrdtExtra}<br />" _
										& "<span class=""SalesPrice"">Our Price: {AttrPrice}</span>" _
										& "</div>" _
										& "</label>" _
										& "</div>"

						pstrOptionTemplate = "<tr>" _
										& "<td {SelectedStyle} valign=""top""><input type=""radio"" name=""{radioName}"" id=""{radioID}"" {OnChangeText} value=""{attrdtId}""{checked}></td>" _
										& "<td {SelectedStyle} nowrap><label for=""{radioID}"">{AttrTitle}<br />" _
										& "Rec. Retail: {attrdtExtra}&nbsp;&nbsp;<br />" _
										& "<span class=""SalesPrice"">Our Price: {AttrPrice}</span>" _
										& "</label>" _
										& "</td>" _
										& "</tr>"

						If pblnDefault Then
							pstrOptionLineOut = Replace(pstrOptionTemplate, "{SelectedStyle}", " class=""attributeDisplaySelected""")
						Else
							pstrOptionLineOut = Replace(pstrOptionTemplate, "{SelectedStyle}", " class=""attributeDisplay""")
						End If
						
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{radioName}", pstrattrName)
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{radioID}", pstrattrName & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID))
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{OnChangeText}", pstrOnChangeText)
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{attrdtId}", paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID))
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{AttrTitle}", pstrAttrTitle)
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{attributeLink}", attributeLink_new(paryAttributes(i)(enAttribute_DetailArray)(j), pstrAttrTitle, False))
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{AttrPrice}", FormatCurrency(pdblAttrDelta, 2))
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{attrdtExtra}", FormatCurrency(paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_Extra), 2))
						pstrOptionLineOut = Replace(pstrOptionLineOut, "{Counter}", iCounter)

						If iCounter > 0 Then
							'pstrOptionLineOut = "<tr><td class=""attributeDisplay"" colspan=""2""><hr /></td></tr>" & pstrOptionLineOut
							pstrOptionLineOut = "<tr class=""attributeDisplay""><td></td><td><hr /></td></tr>" & pstrOptionLineOut
						End If

						If pblnChecked Then
							'pobjStringBuilder.Append "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
							pobjStringBuilder.Append Replace(pstrOptionLineOut, "{checked}", "")
						Else
							'pobjStringBuilder.Append "<div " & pstrSelectedStyle & "><input type='radio' name='" & pstrattrName & "' " & pstrOnChangeText & " value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """ checked>" & attributeLink(rsProdAttributes, pstrAttrTitle, False) & strAttrPrice & "</div>" & vbcrlf
							pobjStringBuilder.Append Replace(pstrOptionLineOut, "{checked}", " checked")
							pblnChecked = True
						End If
				Case enAttrDisplay_Text:	'text - required
						pobjStringBuilder.Append "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
				Case enAttrDisplay_TextOpt:	'text - optional
						pobjStringBuilder.Append "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
						'mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "setAttributeOptional(""" & strFormName & """, """ & pstrFieldName & """, """ & Server.HTMLEncode(pstrAttrTitle) & """);" & vbcrlf

						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = true;" & vbcrlf
				Case enAttrDisplay_Textarea:	'textarea - required
						pobjStringBuilder.Append "<textarea rows='2' columns='40' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''></textarea>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
				Case enAttrDisplay_TextareaOpt:	'textarea - optional
						pobjStringBuilder.Append "<textarea rows='2' columns='40' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''></textarea>" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = 'true';" & vbcrlf
				Case enAttrDisplay_Checkbox:	'checkbox
						pobjStringBuilder.Append "<input type='checkbox' name='" & pstrFieldName & "' value='on'>" & pstrAttrTitle & strAttrPrice & "<br>" & vbcrlf
				Case enAttrDisplay_SelectShowPrice:	'select (show price)
						If pblnChecked Then
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & pstrAttrTitle & pstrAttrDisplayAlt & "</option>" & vbcrlf
						Else
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """ selected>" & pstrAttrTitle & pstrAttrDisplayAlt & "</option>" & vbcrlf
							pblnChecked = True
						End If
				Case enAttrDisplay_SelectChangeImage:	'custom example
						If pblnChecked Then
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
						Else
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """ selected>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
							pblnChecked = True
						End If
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "maryAttrImages['" & EscapeFormName(strFormName) & "'][" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & "] = " & Chr(34) & AttributeImageName(rsProdAttributes, EscapeFormName(strFormName)) & Chr(34) & ";" & vbcrlf
				Case enAttrDisplay_AttrAsQtyBox:	'attribute as quanity box - assumes MPO installed
						pobjStringBuilder.Append "<tr><td>" & pstrAttrTitle & "</td>" _
										& "<td>" _
										& " <input type=hidden name='" & pstrattrName & "' value='" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & "'>" _
										& " <input type=text name='QUANTITY' size=4 onblur='return validQuantity(this);'>" _
										& "</td></tr>" & vbcrlf
				Case enAttrDisplay_Custom:	'custom example
						pobjStringBuilder.Append "<input type='text' id='" & pstrFieldName & "' name='" & pstrFieldName & "' title=" & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & " value=''>" & vbcrlf
						pobjStringBuilder.Append "<a href=""#"" onMouseOut=""MM_swapImgRestore()""  onMouseOver=""MM_swapImage('button_r1_c1','','images/button_r1_c1_f2.gif',1);"" ><img name=""button_r1_c1"" src=""images/button_r1_c1.gif"" width=""187"" height=""43"" border=""0"" onMouseDown=""MM_openBrWindow('/charm_directorynewmb.asp','','width=550,height=500')""></a></td>"
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".title = " & Chr(34) & Server.HTMLEncode(pstrAttrTitle) & Chr(34) & ";" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "document." & strFormName & "." & pstrFieldName & ".optional = 'true';" & vbcrlf
						mstrssAttributeExtenderjsOut = mstrssAttributeExtenderjsOut & "mtxtCharm = document." & strFormName & "." & pstrFieldName & ";" & vbcrlf
				Case Else: 
						If pblnChecked Then
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
						Else
							pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """ selected>" & pstrAttrTitle & strAttrPrice & "</option>" & vbcrlf
							pblnChecked = True
						End If
						'pobjStringBuilder.Append "<option value=""" & paryAttributes(i)(enAttribute_DetailArray)(j)(enAttributeDetail_ID) & """>" & pstrAttrTitle & strAttrPrice & "</option>"
			End Select

		Next 'j
		
		'Close the attribute display
		If Len(pstrClosingTag) > 0 Then	pobjStringBuilder.Append pstrClosingTag & vbcrlf
		
		If Len(pstrAttrPricing) > 0 Then
			pobjStringBuilder.Append vbcrlf _
							& "<tr><td><script language='javascript'  type='text/javascript'>" & vbcrlf _
							& pstrAttrPricing _
							& "</script></td></tr>"
		End If

	Next 'i
	End If	'isArray(paryAttributes)

	strOut = pobjStringBuilder.concat
	Set pobjStringBuilder = Nothing

End Sub	'DisplayAttributes_New

'**********************************************************************************************************

Function EscapeFormName(ByVal strFormName)
	EscapeFormName = Replace(strFormName, "frm", "")
End Function	'EscapeFormName

'**********************************************************************************************************

Function MakeFormNameSafe(strTitle)

Dim pstrTemp

	pstrTemp = Trim(strTitle & "")
	pstrTemp = "frm" & Replace(pstrTemp,"-","")
	pstrTemp = "frm" & Replace(pstrTemp," ","")

	MakeFormNameSafe = pstrTemp

End Function	'MakeFormNameSafe

'**********************************************************************************************************

Function WriteJavaScript(strScript)

	WriteJavaScript = "<script language='javascript'  type='text/javascript'>" & vbcrlf _
								& strScript & vbcrlf _
								& "</script>" & vbcrlf
	
End Function	'WriteJavaScript

'**********************************************************************************************************

Function WriteTitle(strTitle)

Dim pstrTemp

	pstrTemp = strTitle
	pstrTemp = Replace(pstrTemp,"å","")
	pstrTemp = Replace(pstrTemp,"æ","")
	pstrTemp = Replace(pstrTemp,"ø","")

	WriteTitle = pstrTemp

End Function	'WriteTitle

'**********************************************************************************************************

'**********************************************************************************************************
'
'	Attribute Templates
'
'**********************************************************************************************************

Dim maryAttributeTemplates

'**********************************************************************************************************

Function attributeTemplateHeader(byVal lngTemplateID)
	If Not isArray(maryAttributeTemplates) Then Call LoadAttributeTemplates
	attributeTemplateHeader = maryAttributeTemplates(lngTemplateID)(1)
End Function	'attributeTemplateHeader

Function attributeTemplateOddRow(byVal lngTemplateID)
	If Not isArray(maryAttributeTemplates) Then Call LoadAttributeTemplates
	attributeTemplateOddRow = maryAttributeTemplates(lngTemplateID)(2)
End Function	'attributeTemplateOddRow

Function attributeTemplateEvenRow(byVal lngTemplateID)
	If Not isArray(maryAttributeTemplates) Then Call LoadAttributeTemplates
	If Len(maryAttributeTemplates(lngTemplateID)(3)) = 0 Then
		attributeTemplateEvenRow = attributeTemplateOddRow(lngTemplateID)
	Else
		attributeTemplateEvenRow = maryAttributeTemplates(lngTemplateID)(3)
	End If
End Function	'attributeTemplateEvenRow

Function attributeTemplateSelectedRow(byVal lngTemplateID, byVal blnEven)
	If Not isArray(maryAttributeTemplates) Then Call LoadAttributeTemplates
	If Len(maryAttributeTemplates(lngTemplateID)(4)) = 0 Then
		If blnEven Then
			attributeTemplateSelectedRow = attributeTemplateEvenRow(lngTemplateID)
		Else
			attributeTemplateSelectedRow = attributeTemplateOddRow(lngTemplateID)
		End If
	Else
		attributeTemplateSelectedRow = maryAttributeTemplates(lngTemplateID)(4)
	End If
End Function	'attributeTemplateSelectedRow

Function attributeTemplateFooter(byVal lngTemplateID)
	If Not isArray(maryAttributeTemplates) Then Call LoadAttributeTemplates
	attributeTemplateFooter = maryAttributeTemplates(lngTemplateID)(5)
End Function	'attributeTemplateFooter

'**********************************************************************************************************

Sub LoadAttributeTemplates

Dim paryAttributeTemplate(5)
'0 - Template Name
'1 - Start new attribute
'2 - repeating item - odd
'3 - repeating item - even; defaults to odd if not present
'4 - repeating item - selected; defaults to odd/even if not present
'5 - Close attribute

'available items:
'Attribute
'- Attribute Name
'- URL
'- Extra
'- Display Style

'Attribute detail
'- Name
'- URL
'- Extra
'- Image Path
'- Selected
'- Price
'- PriceType
'- Weight

'Example
'Opening
'<tr><td align='left'>{attributeDisplayText}</td>
'<td><select style="{attributeStyle}" name='{attributeName}' id='{attributeName}'>
'Repeating
'<option style="{attributeDetailStyle}" value="{attributeDetailValue}">{attributeDetailDisplayText}<option>
'Closing
'</select></td></tr>

	ReDim maryAttributeTemplates(0)
	
	paryAttributeTemplate(0) = "Standard Select"
	paryAttributeTemplate(1) = "<tr><td align='left'>{attributeDisplayText}</td><td><select style=""{attributeStyle}"" name=""{attributeName}"" id=""{attributeName}"">"
	paryAttributeTemplate(2) = "<option style=""{attributeDetailStyle}"" value=""{attributeDetailValue}"" {attributeDetailSelected}>{attributeDetailDisplayText}<option>"
	paryAttributeTemplate(3) = ""
	paryAttributeTemplate(4) = ""
	paryAttributeTemplate(5) = "</select></td></tr>"
	maryAttributeTemplates(0) = paryAttributeTemplate

End Sub	'LoadAttributeTemplates
%>
