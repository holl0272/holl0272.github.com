<%Option Explicit
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

Response.Buffer = True

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

'	NONE

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Global variables
'**********************************************************

Const cblnOrdersToXMLDebug = False	'True	False

'**********************************************************
'*	Functions
'**********************************************************

'Function CreateItemXML(byVal strTableName, byVal strKeyFieldName, byVal blnKeyIsNumeric, byVal strKeys)
'Function exportItems(byRef strOrderIDs, byVal strXSLFilePath)
'Sub TestWriteXML(byRef objXMLDoc)
'Function WriteXSL(byRef objXML, byVal strXSLFilePath)

'***********************************************************************************************
'***********************************************************************************************

Function CreateItemXML(byVal strTableName, byVal strKeyFieldName, byVal blnKeyIsNumeric, byVal strKeys)

Dim fieldCounter
Dim pobjRS
Dim pstrFieldName
Dim pstrFieldValue
Dim pstrSQL
Dim xmlRoot
Dim xmlNode
Dim xmlProduct
Dim xmlDoc
Dim pstrPrevID

	Set xmlDoc = server.CreateObject("MSXML2.DOMDocument.3.0")
	' Create processing instruction and document root
    Set xmlNode = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'")
    Set xmlNode = xmlDoc.insertBefore(xmlNode, xmlDoc.childNodes.Item(0))
   
	' Create document root
    Set xmlRoot = xmlDoc.createElement("items")
    Set xmlDoc.documentElement = xmlRoot
    xmlRoot.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"

	If Len(strKeys) > 0 Then
		If Right(strKeys, 1) = "," Then strKeys = Left(strKeys, Len(strKeys) - 1)
		strKeys = Replace(strKeys, ", ", ",")
	End If
	
	If blnKeyIsNumeric Then
		pstrSQL = "SELECT * From " & strTableName & " WHERE " & strKeyFieldName & " In (" & strKeys & ")"
	Else
		pstrSQL = "SELECT * From " & strTableName & " WHERE " & strKeyFieldName & " In ('" & Replace(strKeys, ",", "','") & "')"
	End If
			
	Set pobjRS  = GetRS(pstrSQL)
	With pobjRS
		Do While Not .EOF
	
			'Create item Node
			Set xmlProduct = xmlDoc.createElement("item")
			xmlRoot.appendChild xmlProduct
			pstrPrevID = Trim(.Fields(strKeyFieldName).Value)
			xmlProduct.setAttribute "id", pstrPrevID

			'add the root product elements
			For fieldCounter = 1 To .Fields.Count
				pstrFieldName = Trim(.Fields(fieldCounter-1).Name & "")
				pstrFieldValue = getRSFieldValue_Unknown(.Fields(fieldCounter-1))
				
				'remove carriage returns
				'pstrFieldValue = Replace(pstrFieldValue, vbcrlf, "")
				Select Case .Fields(fieldCounter-1).Type
					Case 1:
						Call addNode(xmlDoc, xmlProduct, pstrFieldName, pstrFieldValue)
					Case 2:
						Call addCData(xmlDoc, xmlProduct, pstrFieldName, pstrFieldValue)
					Case Else
						Call addCData(xmlDoc, xmlProduct, pstrFieldName, pstrFieldValue)
				End Select

			Next 'fieldCounter

			.MoveNext
		Loop
		.Close
	End With	'	pobjRS	
	Set pobjRS = Nothing
	
	Set CreateItemXML = xmlDoc
	
	If cblnOrdersToXMLDebug Then
		xmlDoc.preserveWhiteSpace = False
		Response.Write "<fieldset><legend>CreateItemXML</legend><textarea rows=80 cols=120>" & vbcrlf & xmlDoc.xml & vbcrlf & "</textarea></fieldset>"
	End If

End Function	'CreateItemXML

'***************************************************************************************************************************************************************

Function exportItems(byVal strTableName, byVal strKeyFieldName, byVal blnKeyIsNumeric, byVal strKeys, byVal strXSLFilePath)

Dim objXML
Dim pstrBody

'	On Error Resume Next

	If Len(strKeys) = 0 Then Exit Function
		
	Set objXML = CreateItemXML(strTableName, strKeyFieldName, blnKeyIsNumeric, strKeys)
	pstrBody = WriteXSL(objXML, strXSLFilePath)
	pstrBody = Replace(pstrBody,"<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-16" & Chr(34) & "?>", "")

	'Response.Write pstrBody
	exportItems = pstrBody

End Function	'exportItems

'***************************************************************************************************************************************************************

Sub SendOrdersXMLToResponse(byRef strOrderIDs)

Dim objXML

'	On Error Resume Next

	If Len(strOrderIDs) = 0 Then Exit Sub
	'If Right(strOrderIDs, 1) = "," Then strOrderIDs = Left(strOrderIDs, Len(strOrderIDs) - 1)	'check added because sometimes a stray comma appears
		
	If LoadOrdersXML(strOrderIDs, objXML) Then
		Response.ContentType = "text/xml"
		objXML.Save Response
		'Call debug.SaveToDisk("db\text.xml", objXML.XML, True)
		Set objXML = Nothing
	End If

End Sub	'SendOrdersXMLToResponse

'***********************************************************************************************

Sub TestWriteXML(byRef objXMLDoc)
	If cblnOrdersToXMLDebug Then	
		Response.Write objXMLDoc.xml
	End If
End Sub	'TestWriteXML

'***************************************************************************************************************************************************************

Function WriteXSL(byRef objXML, byVal strXSLFilePath)

Dim objXSL
Dim strOutput

	' Load the XSL from the XSL file
	set objXSL = Server.CreateObject("MSXML2.DOMDocument")
	objXSL.async = false
	'objXSL.preserveWhiteSpace = True
	'debugprint "strXSLFilePath", strXSLFilePath
	If objXSL.Load(strXSLFilePath) Then
		If cblnOrdersToXMLDebug Then
			Response.Write "<fieldset><legend>WriteXSL - XSL</legend>"
			Call TestWriteXML(objXSL)
			Response.Write "</fieldset>"
		End If
		strOutput = objXML.transformNode(objXSL)
	Else
		strOutput = "Error Loading XSL document " & strXSLFilePath & "."
	End If
	Set objXML = Nothing
	
	strOutput = Replace(strOutput,"&amp;nbsp;","&nbsp;")
	WriteXSL = strOutput

End Function	'WriteXSL

'***************************************************************************************************************************************************************

'--------------------------------------------------------------------------------------------------
%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="ssLibrary/clsDebug.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
  
On Error Goto 0	'added because of global error suppression in mail.asp

'**************************************************
'
'	Start Code Execution
'

'page variables
Dim maryTemp
Dim mstrAction
Dim mstrExportField
Dim mstrXSLFilePath

Dim mstrTableName
Dim mstrTableKeyFieldName
Dim mstrTableKeyFieldIsNumeric
Dim mstrExportTemplateDirectory
Dim mstrIDsToExport

Const filename = "Export.csv"

    maryTemp = Split(LoadRequestValue("Action"), "|")
    
	mstrAction = maryTemp(0)
	If UBound(maryTemp) > 0 Then mstrXSLFilePath = Trim(maryTemp(1))
	If UBound(maryTemp) > 1 Then mstrExportField = Trim(maryTemp(2))
	
	mstrTableName = LoadRequestValue("tableName")
	mstrTableKeyFieldName = LoadRequestValue("tableKeyFieldName")
	mstrTableKeyFieldIsNumeric = LoadRequestValue("tableKeyFieldIsNumeric")
	mstrExportTemplateDirectory = LoadRequestValue("ExportTemplateDirectory")
	mstrIDsToExport = LoadRequestValue("chkItemID")

	If cblnOrdersToXMLDebug Then
		Response.Write "<fieldset><legend>Key Parameters</legend>"
		Response.Write "mstrXSLFilePath: " & mstrXSLFilePath & "<BR>"
		Response.Write "mstrExportField: " & mstrExportField & "<BR>"
		Response.Write "tableName: " & mstrTableName & "<BR>"
		Response.Write "tableKeyFieldName: " & mstrTableKeyFieldName & "<BR>"
		Response.Write "tableKeyFieldIsNumeric: " & mstrTableKeyFieldIsNumeric & "<BR>"
		Response.Write "ExportTemplateDirectory: " & mstrExportTemplateDirectory & "<BR>"
		Response.Write "IDs To Export: " & mstrIDsToExport & "<BR>"
		Response.Write "</fieldset>"
	End If
	
    Select Case mstrAction
        Case "downloadOrders"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrExportTemplateDirectory & mstrXSLFilePath)
			Response.ContentType = "application/octet-stream"
			Response.AddHeader "Content-Disposition", "attachment; filename=""" & filename & """"
			Response.Write exportItems(mstrTableName, mstrTableKeyFieldName, mstrTableKeyFieldIsNumeric, mstrIDsToExport, mstrXSLFilePath)
        Case "printOrders"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrExportTemplateDirectory & mstrXSLFilePath)
			Call exportItems(mstrTableName, mstrTableKeyFieldName, mstrTableKeyFieldIsNumeric, mstrIDsToExport, mstrXSLFilePath)
			Response.Write "<OBJECT ID=WebBrowser1 WIDTH=0 HEIGHT=0 CLASSID='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2'></OBJECT>"
			Response.Write "<script language=javascript>" _
						   & "document.all('WebBrowser1').ExecWB(6, 2);window.close();</script>"
        Case "viewOrders"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrExportTemplateDirectory & mstrXSLFilePath)
			Response.Write exportItems(mstrTableName, mstrTableKeyFieldName, mstrTableKeyFieldIsNumeric, mstrIDsToExport, mstrXSLFilePath)
    End Select
    
    Call ReleaseObject(cnn)
   
%>