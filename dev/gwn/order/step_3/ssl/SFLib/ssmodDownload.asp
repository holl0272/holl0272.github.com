<%
'********************************************************************************
'*   Sandshot Software Product Download Page                                    *
'*   Release Version   1.0	                                                    *
'*   Release Date      November 16, 2002										*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2006 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'***********************************************************************************************
'	Page level variables
'***********************************************************************************************

Const cblnDisableDownloadableProducts = True

'Display error enumerations
Const enDownloadRequest_Valid = 0
Const enDownloadRequest_InValidCustomerID = 1
Const enDownloadRequest_InValidOrderDetailID = 2
Const enDownloadRequest_PastDownloadExpiration = 3
Const enDownloadRequest_DownloadCountReached = 4
Const enDownloadRequest_InvalidRequest = 5
Const enDownloadRequest_NoDownloadAvailable = 6
Const enDownloadRequest_NotAuthorized = 7
Const enDownloadRequest_InvalidFilePath = 8

Dim Download_ExpiresOn
Dim Download_LastDownload
Dim Download_CurrentDownloadCount
Dim Download_MaxDownloads
Dim Download_FilePath
Dim Download_FileName
Dim Download_FileSize
Dim Download_RequestStatus

'***********************************************************************************************
'	Functions
'***********************************************************************************************

'Function ConvertFilePath(byVal strFilePath)
'Function validateDownloadFileAvailability(byVal lngOrderDetailID, byRef enDownloadRequest, byRef strFileName, byRef strFilePath)
'Function HasDownloadAvailable_orderDetail(byVal lngCustomerID, byVal lngOrderDetailID)
'Function InitiateDownload(byVal strFilePath)
'Function NumDownloadsByOrderDetailID(byVal lngOrderDetailID)
'Sub RecordDownloadCompleted(byVal ssFileDownloadID, byVal ssFileDownloadInitiated, byVal ssFileDownloadOrderItemID, byVal ssFileDownloadedFileName, byVal ssFileDownloadedVersion)

'***********************************************************************************************
'	Functions
'***********************************************************************************************

Sub initializeDownloadVariables

	Download_ExpiresOn = ""
	Download_LastDownload = ""
	Download_CurrentDownloadCount = 0
	Download_MaxDownloads = 0
	Download_FilePath = ""
	Download_FileName = ""
	Download_FileSize = 0
	Download_RequestStatus = enDownloadRequest_NoDownloadAvailable

End Sub	'initializeDownloadVariables

'***********************************************************************************************

Function DownloadRequest_RequestStatusText(byVal bytDownload_RequestStatus)

Dim pstrOut

	Select Case bytDownload_RequestStatus
		Case enDownloadRequest_Valid:					pstrOut = "Valid"
		Case enDownloadRequest_InValidCustomerID:		pstrOut = "Invalid Customer Number"
		Case enDownloadRequest_InValidOrderDetailID:	pstrOut = "Invalid Order Identifier"
		Case enDownloadRequest_PastDownloadExpiration:	pstrOut = "Download Period Has Expired"
		Case enDownloadRequest_DownloadCountReached:	pstrOut = "Maximum Download Limit Reached"
		Case enDownloadRequest_InvalidRequest:			pstrOut = "Invalid Request"
		Case enDownloadRequest_NoDownloadAvailable:		pstrOut = "No Download Available"
		Case enDownloadRequest_NotAuthorized:			pstrOut = "Not Authorized"
		Case enDownloadRequest_InvalidFilePath:			pstrOut = "Invalid File Path"
		Case Else:										pstrOut = "Unexpected Error"
	End Select

	DownloadRequest_RequestStatusText = pstrOut

End Function	'DownloadRequest_RequestStatusText

'***********************************************************************************************

Function hasRemainingDownloads()

Dim pblnResult

	If Download_MaxDownloads = 0 Then
		pblnResult = True
	Else
		pblnResult = CBool(Download_MaxDownloads > Download_CurrentDownloadCount)
	End If
	
	If pblnResult Then
		If Len(Download_ExpiresOn) = 0 Then
			pblnResult = True
		ElseIf Download_ExpiresOn < Now() Then
			pblnResult = False
		End If
	End If

	hasRemainingDownloads = pblnResult

End Function	'hasRemainingDownloads

'***********************************************************************************************

Function ConvertFilePath(byVal strFilePath)

Dim pstrFilePathOut
Dim pstrRoot
Dim pstrWebRoot

	pstrFilePathOut = strFilePath
	pstrWebRoot = Request.ServerVariables("APPL_PHYSICAL_PATH")
	pstrRoot = Replace(pstrWebRoot, "wwwroot\", "")
	
	pstrFilePathOut = Replace(pstrFilePathOut, "{webRoot}", pstrWebRoot)
	pstrFilePathOut = Replace(pstrFilePathOut, "{root}", pstrRoot)
	
	'Response.Write "pstrFilePathOut: " & pstrFilePathOut & "<br />" & vbcrlf

	ConvertFilePath = pstrFilePathOut

End Function	'ConvertFilePath

'***********************************************************************************************

Function DownloadFile(byVal strFilePath, byVal lngOrderDetailID)

Dim plngID
Dim pstrResult

	plngID = RecordDownloadInitiated(lngOrderDetailID, strFilePath, "")
	If plngID <> -1 Then
		If InitiateDownload(strFilePath) Then
			Call RecordDownloadCompleted(plngID)
		Else
			pstrResult = "File Download Error: Please contact customer service with the following code <b>InitiateDownload." & mlngcustID & "-" & mlngOrderDetailID & "</b>"
		End If
	Else
		pstrResult = "File Download Error: Please contact customer service with the following code <b>RecordDownloadInitiated." & mlngcustID & "-" & mlngOrderDetailID & "</b>"
	End If
	
	DownloadFile = pstrResult

End Function	'DownloadFile

'***********************************************************************************************

Function validateDownloadFileAvailability(byVal lngOrderDetailID, byRef enDownloadRequest, byRef strFileName, byRef strFilePath, byRef strFileSize)

Dim i
Dim pblnSuccess
Dim plngPos
Dim pobjCMDFileFragement
Dim	pobjCMDOrderDetail
Dim pobjFileFragement
Dim pobjRS
Dim pobjRSOrderDetails
Dim pstrFilePath
Dim pstrAttributeCategory
Dim pstrAttributeDetail
Dim pstrFileFragment
Dim pstrFileName
Dim pstrProductID
Dim pstrSQL

	pblnSuccess = True
	pstrFilePath = ConvertFilePath(strFilePath)

	If Len(lngOrderDetailID) = 0 Or Not isNumeric(lngOrderDetailID) Then
		enDownloadRequest = enDownloadRequest_InValidOrderDetailID
		lngOrderDetailID = ""
		pblnSuccess = False
	ElseIf Len(pstrFilePath) = 0 Then
		enDownloadRequest = enDownloadRequest_NoDownloadAvailable
		pblnSuccess = False
	Else
		If InStr(1, strFilePath, "{") > 0 Then
			'Replacement codes in path so need to check attributes

			If False Then
				'this section only used for debugging
				'split the path by {
				Dim paryAttributeParameters_temp
				Dim paryAttributeParameters
				paryAttributeParameters_temp = Split(strFilePath, "{")
				
				'remove everything after the }
				For i = 0 To UBound(paryAttributeParameters_temp)
					plngPos = InStr(paryAttributeParameters_temp(i), "}")
					If plngPos > 0 Then
						paryAttributeParameters_temp(i) = Left(paryAttributeParameters_temp(i), plngPos-1)
					Else
						paryAttributeParameters_temp(i) = ""
					End If
				Next 'i
				
				'set attributes to working array
				ReDim paryAttributeParameters(UBound(paryAttributeParameters_temp) - 1, 1)
				For i = 1 To UBound(paryAttributeParameters_temp)
					paryAttributeParameters(i-1, 0) = paryAttributeParameters_temp(i)
				Next 'i
				
				Response.Write "<fieldset><legend>paryAttributeParameters</legend>"
				For i = 0 To UBound(paryAttributeParameters)
					Response.Write i & ": " & paryAttributeParameters(i, 0) & "<br />"
				Next 'i
				Response.Write "</fieldset>"
			End If	'False

			pstrSQL = "SELECT sfOrderDetails.odrdtProductID, sfOrderAttributes.odrattrAttribute, sfOrderAttributes.odrattrName" _
					& " FROM sfOrderDetails LEFT JOIN sfOrderAttributes ON sfOrderDetails.odrdtID = sfOrderAttributes.odrattrOrderDetailId" _
					& " WHERE sfOrderDetails.odrdtID=?"
			Set pobjCMDOrderDetail = CreateObject("ADODB.Command")
			pobjCMDOrderDetail.ActiveConnection = cnn
			pobjCMDOrderDetail.CommandType = adCmdText
			pobjCMDOrderDetail.CommandText = pstrSQL

			pobjCMDOrderDetail.Parameters.Append pobjCMDOrderDetail.CreateParameter("odrdtID", adInteger, adParamInput, 4, lngOrderDetailID)

			On Error Resume Next
			Set	pobjRSOrderDetails = pobjCMDOrderDetail.Execute
			If Err.number <> 0 Then
				If ssDebug_Download Then Response.Write "<font color=red>Error in validateDownloadFileAvailability: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				If ssDebug_Download Then Response.Write "<font color=red>Error in validateDownloadFileAvailability: sql = " & pstrSQL & "</font><br />" & vbcrlf
				Err.Clear
				pblnSuccess = False
			ElseIf Not pobjRSOrderDetails.EOF Then
				On Error Goto 0

				pstrSQL = "SELECT sfAttributeDetail.attrdtFileName" _
						& " FROM (sfProducts INNER JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) INNER JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
						& " WHERE ((sfAttributeDetail.attrdtName=?) AND (sfAttributes.attrName=?) AND (sfProducts.prodID=?))"

				Do While Not pobjRSOrderDetails.EOF
					pstrProductID = Trim(pobjRSOrderDetails.Fields("odrdtProductID").Value & "")
					pstrAttributeCategory = Trim(pobjRSOrderDetails.Fields("odrattrName").Value & "")
					pstrAttributeDetail = Trim(pobjRSOrderDetails.Fields("odrattrAttribute").Value & "")
					
					'Now adjust for attribute extender which MAY save the category name: in the detail
					If inStr(1, pstrAttributeDetail, pstrAttributeCategory & ": ") = 1 Then pstrAttributeDetail = Replace(pstrAttributeDetail,  pstrAttributeCategory & ": ", "")
					
					If False Then
						pstrSQL = "SELECT sfAttributeDetail.attrdtFileName" _
								& " FROM (sfProducts INNER JOIN sfAttributes ON sfProducts.prodID = sfAttributes.attrProdId) INNER JOIN sfAttributeDetail ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId" _
								& " WHERE ((sfAttributeDetail.attrdtName='" & Replace(pstrAttributeDetail, "'", "''") & "') AND (sfAttributes.attrName='" & Replace(pstrAttributeCategory, "'", "''") & "') AND (sfProducts.prodID='" & Replace(pstrProductID, "'", "''") & "'))"
						set	pobjFileFragement = CreateObject("adodb.recordset")
						pobjFileFragement.CursorLocation = 2 'adUseClient
					    
						On Error Resume Next
						pobjFileFragement.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
					Else
						If isObject(pobjCMDFileFragement) Then
							With pobjCMDFileFragement
								.Parameters(0).Value = pstrAttributeDetail
								.Parameters(1).Value = pstrAttributeCategory
								.Parameters(2).Value = pstrProductID
							End With	'pobjCMDFileFragement
						Else
							Set pobjCMDFileFragement = CreateObject("ADODB.Command")
							With pobjCMDFileFragement
								.ActiveConnection = cnn
								.CommandType = adCmdText
								.CommandText = pstrSQL

								'.Parameters.Append .CreateParameter("attrdtFileName", adVarChar, adParamInputOutput, 255, NULL)
								.Parameters.Append .CreateParameter("attrdtName", adVarChar, adParamInput, 255, pstrAttributeDetail)
								.Parameters.Append .CreateParameter("attrName", adVarChar, adParamInput, 255, pstrAttributeCategory)
								.Parameters.Append .CreateParameter("prodID", adVarChar, adParamInput, 255, pstrProductID)

							End With	'pobjCMDFileFragement
							'Set pobjCMDFileFragement = Nothing
						End If

						On Error Resume Next
						Set	pobjFileFragement = pobjCMDFileFragement.Execute
							
					End If	'used because Command object not working

					If Err.number <> 0 Then
						If ssDebug_Download Then Response.Write "<font color=red>Error in validateDownloadFileAvailability: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
						If ssDebug_Download Then Response.Write "<font color=red>Error in validateDownloadFileAvailability: sql = " & pstrSQL & "</font><br />" & vbcrlf
						Err.Clear
						pblnSuccess = False
					ElseIf Not pobjFileFragement.EOF Then
						pstrFileFragment = Trim(pobjFileFragement.Fields("attrdtFileName").Value & "")
						If ssDebug_Download Then Response.Write "<b>pstrFileFragment: " & pstrFileFragment & "</b><br />"
					Else
						If ssDebug_Download Then
							Response.Write "<fieldset><legend>No Match Found</legend>"
							Response.Write "<font color=red>No Match Found</font><br />"
							Response.Write "Product ID: " & pstrProductID & "<br />"
							Response.Write "Attribute Category: " & pstrAttributeCategory & "<br />"
							Response.Write "Attribute Detail: " & pstrAttributeDetail & "<hr>"
							Response.Write "sql = " & pstrSQL & ""
							Response.Write "</fieldset>"
						End If
					End If
					pobjFileFragement.Close
					Set	pobjFileFragement = Nothing
					
					pstrFilePath = Replace(pstrFilePath, "{" & pstrAttributeCategory & "}", pstrFileFragment)

					pobjRSOrderDetails.MoveNext
				Loop
				pobjRSOrderDetails.Close

				If isObject(pobjCMDFileFragement) Then
					Set pobjCMDFileFragement = Nothing
				End If

			End If	'Err.number <> 0
			Set	pobjRSOrderDetails = Nothing
			Set pobjCMDOrderDetail = Nothing

		End If	'InStr(1, strFilePath, "{") < 1

	End If	'Len(lngOrderDetailID) = 0 Or Not isNumeric(lngOrderDetailID)

	If ssDebug_Download Then Response.Write "<b>Final pstrFilePath: " & pstrFilePath & "</b><br />"

	'Now set the filename
	Dim pobjFSO
	Dim pobjFile
	strFilePath = pstrFilePath
	
	Set pobjFSO = CreateObject("Scripting.FileSystemObject")
	If pobjFSO.FileExists(strFilePath) Then
		Set pobjFile = pobjFSO.GetFile(strFilePath)
		strFileName = pobjFile.Name
		strFileSize = pobjFile.Size
		Set pobjFile = Nothing

		enDownloadRequest = enDownloadRequest_Valid
	Else
		pblnSuccess = False
		enDownloadRequest = enDownloadRequest_InvalidFilePath
	End If
	Set pobjFSO = Nothing
	
	'plngPos = InStrRev(strFilePath, "\")
	'If plngPos < Len(strFilePath) Then strFileName = Right(strFilePath, Len(strFilePath) - plngPos)
	
	validateDownloadFileAvailability = pblnSuccess
	
End Function	'validateDownloadFileAvailability

'**********************************************************************************************************

Function HasDownloadAvailable_orderDetail(byVal lngCustomerID, byVal lngOrderDetailID)

Dim pblnSuccess
Dim pobjCMD
Dim pobjRS
Dim pstrFilePath
Dim pstrFileName
Dim pstrSQL

    If cblnDisableDownloadableProducts Then
        HasDownloadAvailable_orderDetail = False
		Download_RequestStatus = enDownloadRequest_NoDownloadAvailable
        Exit Function
    End If
    
	pblnSuccess = True
	Call initializeDownloadVariables
	
	If Len(lngCustomerID) = 0 Or Not isNumeric(lngCustomerID) Then
		Download_RequestStatus = enDownloadRequest_InValidCustomerID
		lngCustomerID = ""
		pblnSuccess = False
	End If
	If Len(lngOrderDetailID) = 0 Or Not isNumeric(lngOrderDetailID) Then
		Download_RequestStatus = enDownloadRequest_InValidOrderDetailID
		lngOrderDetailID = ""
		pblnSuccess = False
	End If

	If pblnSuccess Then
		pstrSQL = "SELECT sfOrderDetails.odrdtDownloadExpiresOn, sfOrderDetails.odrdtMaxDownloads, sfOrderDetails.odrdtDownloadAuthorized, sfProducts.prodFileName, Count(ssFileDownloads.ssFileDownloadCompleted) AS CountOfssFileDownloadCompleted" _
				& " FROM ssFileDownloads RIGHT JOIN ((sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID) ON ssFileDownloads.ssFileDownloadOrderItemID = sfOrderDetails.odrdtID" _
				& " GROUP BY sfOrderDetails.odrdtDownloadExpiresOn, sfOrderDetails.odrdtMaxDownloads, sfOrderDetails.odrdtDownloadAuthorized, sfProducts.prodFileName, sfOrders.orderID, sfOrders.orderIsComplete, sfOrders.orderCustId, sfOrderDetails.odrdtID" _
				& " HAVING ((sfOrders.orderIsComplete=1) AND (sfOrders.orderCustId=?) AND (sfOrderDetails.odrdtID=?))"

		pstrSQL = "SELECT sfOrderDetails.odrdtDownloadExpiresOn, sfOrderDetails.odrdtMaxDownloads, sfOrderDetails.odrdtDownloadAuthorized, sfProducts.prodFileName, Count(ssFileDownloads.ssFileDownloadCompleted) AS CountOfssFileDownloadCompleted, Max(ssFileDownloads.ssFileDownloadCompleted) AS MaxOfssFileDownloadCompleted" _
				& " FROM ssFileDownloads RIGHT JOIN ((sfOrders LEFT JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) LEFT JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID) ON ssFileDownloads.ssFileDownloadOrderItemID = sfOrderDetails.odrdtID" _
				& " GROUP BY sfOrderDetails.odrdtDownloadExpiresOn, sfOrderDetails.odrdtMaxDownloads, sfOrderDetails.odrdtDownloadAuthorized, sfProducts.prodFileName, sfOrders.orderID, sfOrders.orderIsComplete, sfOrders.orderCustId, sfOrderDetails.odrdtID" _
				& " HAVING ((sfOrders.orderIsComplete=1) AND (sfOrders.orderCustId=?) AND (sfOrderDetails.odrdtID=?))"

		Set pobjCMD = CreateObject("ADODB.Command")
		With pobjCMD
			.ActiveConnection = cnn
			.CommandType = adCmdText
			.CommandText = pstrSQL

			.Parameters.Append .CreateParameter("orderCustId", adInteger, adParamInput, 4, lngCustomerID)
			.Parameters.Append .CreateParameter("sfOrderDetails", adInteger, adParamInput, 4, lngOrderDetailID)

			On Error Resume Next
			Set	pobjRS = .Execute
			If Err.number <> 0 Then
				If ssDebug_Download Then
					Response.Write "<fieldset><legend>Error in HasDownloadAvailable_orderDetail</legend>" _
								& "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" _
								& "pstrSQL: " & pstrSQL & "<br />" _
								& "</fieldset>"

					Err.Clear
				End If
				pblnSuccess = False
			Else
				On Error Goto 0
				With pobjRS
					If .EOF Then
						Download_RequestStatus = enDownloadRequest_InvalidRequest
						pblnSuccess = False
					Else

						Download_LastDownload = Trim(.Fields("MaxOfssFileDownloadCompleted").Value & "")
						Download_ExpiresOn = Trim(.Fields("odrdtDownloadExpiresOn").Value & "")
						
						Download_MaxDownloads = Trim(.Fields("odrdtMaxDownloads").Value & "")
						If Len(Download_MaxDownloads) > 0 Then
							Download_MaxDownloads = CLng(Download_MaxDownloads)
						Else
							Download_MaxDownloads = 0
						End If
						
						Download_CurrentDownloadCount = Trim(.Fields("CountOfssFileDownloadCompleted").Value & "")
						If Len(Download_CurrentDownloadCount) > 0 Then
							Download_CurrentDownloadCount = CLng(Download_CurrentDownloadCount)
						Else
							Download_CurrentDownloadCount = 0
						End If
						
						pstrFilePath = Trim(.Fields("prodFileName").Value & "")
						
						'not valid if no file name
						If Len(pstrFilePath) = 0 Then 
							Download_RequestStatus = enDownloadRequest_NoDownloadAvailable
							pblnSuccess = False
						Else
							'Check for valid path
							If inStr(1, pstrFilePath, ":") = 1 And inStr(1, pstrFilePath, "{root}") <> 1 Then pstrFilePath = getConfigurationSettingFromCache("DownloadRootLocation_Default", "") & pstrFilePath
							If validateDownloadFileAvailability(lngOrderDetailID, Download_RequestStatus, pstrFileName, pstrFilePath, Download_FileSize) Then
								Download_FileName = pstrFileName
								Download_FilePath = pstrFilePath
								
								Download_RequestStatus = enDownloadRequest_Valid
								pblnSuccess = True
								
								'check if authorized
								If ConvertToBoolean(getConfigurationSettingFromCache("Download_CheckForAuthorization", True), False) And pblnSuccess Then
									If Len(Trim(.Fields("odrdtDownloadAuthorized").Value & "")) > 0 Then 
										If .Fields("odrdtDownloadAuthorized").Value = 0 Then
											Download_RequestStatus = enDownloadRequest_NotAuthorized
											pblnSuccess = False
										End If
									Else
										Download_RequestStatus = enDownloadRequest_NotAuthorized
										pblnSuccess = False
									End If
								End If

								'check if past expiration date
								If pblnSuccess And Len(CStr(Download_ExpiresOn)) > 0 Then 
									If CBool(Download_ExpiresOn < Now()) Then
										Download_RequestStatus = enDownloadRequest_PastDownloadExpiration
										pblnSuccess = False
									End If
								End If

								'check if past download limit
								If pblnSuccess And Download_MaxDownloads > 0 Then
									If CBool(Download_MaxDownloads <= Download_CurrentDownloadCount) Then
										Download_RequestStatus = enDownloadRequest_DownloadCountReached
										pblnSuccess = False
									End If
								End If
								
							End If	'validateDownloadFileAvailability
							
						End If	'Len(pstrFilePath) = 0
					
					End If	'Err.number <> 0

					If ssDebug_Download Then	'True	False
						Response.Write "<fieldset><legend>HasDownloadAvailable_orderDetail</legend>" _
									& "pstrSQL: " & pstrSQL & "<br />" _
									& "lngCustomerID: " & lngCustomerID & "<br />" _
									& "lngOrderDetailID: " & lngOrderDetailID & "<hr>" _
									& "odrdtDownloadAuthorized: " & Trim(.Fields("odrdtDownloadAuthorized").Value) & "<br />" _
									& "Download_ExpiresOn: " & Download_ExpiresOn & "<br />" _
									& "Download_LastDownload: " & Download_LastDownload & "<br />" _
									& "Download_CurrentDownloadCount: " & Download_CurrentDownloadCount & "<br />" _
									& "Download_MaxDownloads: " & Download_MaxDownloads & "<br />" _
									& "Download_FileName: " & Download_FileName & "<br />" _
									& "Download_FilePath: " & Download_FilePath & "<br />" _
									& "Download_FileSize: " & Download_FileSize & "<br />" _
									& "Download_RequestStatus: " & DownloadRequest_RequestStatusText(Download_RequestStatus) & "(" & Download_RequestStatus & ")<br />" _
									& "</fieldset>"
					End If
				
					.Close
				End With	'pobjRS
				Set	pobjRS = Nothing

			End If	'Err.number <> 0
			
		End With	'pobjCMD
		
		Set pobjCMD = Nothing
		
	End If	'pblnSuccess

	HasDownloadAvailable_orderDetail = pblnSuccess
	
End Function	'HasDownloadAvailable_orderDetail

'**********************************************************************************************************

Function InitiateDownload(byVal strFilePath)

Dim ContentType
Dim FileSize
Dim pblnSuccess
Dim pblnForceDownload
Dim plngPos
Dim pobjFSO
Dim pobjFile
Dim pobjStream
Dim pstrFileExtension
Dim pstrFileName

	pblnForceDownload = True
	pblnSuccess = False

	If Len(strFilePath) > 0 Then

		Set pobjFSO = CreateObject("Scripting.FileSystemObject")
		If pobjFSO.FileExists(strFilePath) Then
			Set pobjFile = pobjFSO.GetFile(strFilePath)
			pstrFileName = pobjFile.Name
			FileSize = pobjFile.Size
			Set pobjFile = Nothing
			
			pstrFileExtension = LCase(Right(strFilePath, 4))
			
			'Content type defined by RFC 822: See http://www.w3.org/Protocols/rfc1341/4_Content-Type.html
			'Content types can be: application, audio, image, message, multipart, text, video
			'Sub content types can be about anything
		
			Select Case pstrFileExtension
				Case ".asf"
					ContentType = "video/x-ms-asf"
				Case ".asp"
					ContentType = "text/asp"
				Case ".avi"
					ContentType = "video/avi"
				Case ".doc"
					ContentType = "application/msword"
				Case ".gif"
					ContentType = "image/gif"
				Case ".htm", "html"
					ContentType = "text/html"
				Case ".jpg", "jpeg"
					ContentType = "image/jpeg"
				Case ".mp3"
					ContentType = "audio/mpeg3"
				Case ".mpg", "mpeg"
					ContentType = "video/mpeg"
				Case ".pdf"
					ContentType = "application/pdf"
				Case ".rtf"
					ContentType = "application/rtf"
				Case ".wav"
					ContentType = "audio/wav"
				Case ".xls"
					ContentType = "application/vnd.ms-excel"
				Case ".zip"
					ContentType = "application/zip"
				Case Else
					'Handle All Other Files
					ContentType = "application/octet-stream"
			End Select
				
			'-- Get file into stream
			Set pobjStream = CreateObject("ADODB.Stream")
			pobjStream.Open
			pobjStream.Type = 1 'binary

			On Error Resume Next
			pobjStream.LoadFromFile strFilePath
			If Err.number = 0 Then
				On Error Goto 0
				
				If Response.Buffer Then Response.Clear
				'-- send stream To response
				Response.AddHeader "Content-Length", FileSize
				If pblnForceDownload Then
					'-- the attachment parameter will force the download regardless of content type
					Response.AddHeader "Content-Disposition", "attachment; filename=""" & pstrFileName & """"
				Else
					Response.AddHeader "Content-Disposition", "filename=""" & pstrFileName & """"
				End If
				Response.ContentType = ContentType
				Response.Charset = "UTF-8"
				Response.BinaryWrite pobjStream.Read(-1) 'read all
				'-- close the stream
				pobjStream.Close
			
				pblnSuccess = True
			Else
				If ssDebug_Download Then Response.Write "<font color=red>Error in InitiateDownload: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				If ssDebug_Download Then Response.Write "<font color=red>pstrFileName: " & pstrFileName & "</font><br />" & vbcrlf
				Err.Clear
			End If	'Err.number = 0
			Set pobjStream = nothing
		
		Else
			If ssDebug_Download Then Response.Write "<font color=red>InitiateDownload : " & strFilePath & " does not exist!</font><br />" & vbcrlf
			pblnSuccess = False
		End If	'pobjFSO.FileExists(strFilePath)
		Set pobjFSO = Nothing

	End If	'Len(strFilePath) > 0

	InitiateDownload = pblnSuccess
	
End Function	'InitiateDownload

'**********************************************************************************************************

Function NumDownloadsByOrderDetailID(byVal lngOrderDetailID)

Dim plngDownloadCount
Dim pobjRS
Dim pstrSQL
Dim pstrFileName


	plngDownloadCount = 0
	
	If Len(lngOrderDetailID) > 0 And isNumeric(lngOrderDetailID) Then
		pstrSQL = "SELECT Count(ssFileDownloads.ssFileDownloadOrderItemID) AS CountOfssFileDownloadOrderItemID" _
				& " FROM ssFileDownloads" _
				& " Where ((ssFileDownloads.ssFileDownloadOrderItemID=" & lngOrderDetailID & ") AND (ssFileDownloads.ssFileDownloadCompleted Is Not Null))" _
				& " GROUP BY ssFileDownloads.ssFileDownloadCompleted"

		Set	pobjRS = CreateObject("adodb.recordset")
		With pobjRS
			.CursorLocation = 2 'adUseClient
	        
			On Error Resume Next
			.Open pstrSQL, cnn, 3, 1	'adOpenStatic, adLockReadOnly
			If Err.number <> 0 Then
				If ssDebug_Download Then Response.Write "<font color=red>Error in NumDownloadsByOrderDetailID: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
				If ssDebug_Download Then Response.Write "<font color=red>Error in NumDownloadsByOrderDetailID: sql = " & pstrSQL & "</font><br />" & vbcrlf
				Err.Clear
			ElseIf Not .EOF Then
				plngDownloadCount = .Fields("CountOfssFileDownloadOrderItemID").Value
			End If	'Err.number <> 0
			.Close
		End With	'pobjRS
		Set pobjRS = Nothing
	End If	'Len(lngOrderDetailID) > 0 And isNumeric(lngOrderDetailID)

	NumDownloadsByOrderDetailID = plngDownloadCount
	
End Function	'NumDownloadsByOrderDetailID

'**********************************************************************************************************

Function RecordDownloadInitiated(byVal ssFileDownloadOrderItemID, byVal ssFileDownloadedFileName, byVal ssFileDownloadedVersion)
							
Dim plngID
Dim pobjCommand
Dim pobjRS
Dim pTempssFileDownloadInitiated
Dim ssFileDownloadDNSHost
Dim ssFileDownloadREMOTE_ADDR
Dim ssFileDownloadREMOTE_HOST
Dim pstrSQL

	plngID = -1
	
	pTempssFileDownloadInitiated = Now()
	ssFileDownloadREMOTE_ADDR = Request.ServerVariables("REMOTE_ADDR")
	ssFileDownloadREMOTE_HOST = Request.ServerVariables("REMOTE_HOST")
	ssFileDownloadDNSHost = ""

	pstrSQL = "Insert Into ssFileDownloads (ssFileDownloadInitiated, ssFileDownloadOrderItemID, ssFileDownloadREMOTE_ADDR, ssFileDownloadREMOTE_HOST, ssFileDownloadDNSHost, ssFileDownloadedFileName, ssFileDownloadedVersion)" _
			& " Values (?, ?, ?, ?, ?, ?, ?)"

	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.ActiveConnection = cnn
		.CommandText = pstrSQL

		'.Parameters.Append .CreateParameter("ssFileDownloadID", adInteger, adParamInputOutput, 4, NULL)

		.Parameters.Append .CreateParameter("ssFileDownloadInitiated", adDBTimeStamp, adParamInput, 16, pTempssFileDownloadInitiated)
		.Parameters.Append .CreateParameter("ssFileDownloadOrderItemID", adInteger, adParamInput, 4, ssFileDownloadOrderItemID)
		.Parameters.Append .CreateParameter("ssFileDownloadREMOTE_ADDR", adVarChar, adParamInput, 255, ssFileDownloadREMOTE_ADDR)
		.Parameters.Append .CreateParameter("ssFileDownloadREMOTE_HOST", adVarChar, adParamInput, 255, ssFileDownloadREMOTE_HOST)
		.Parameters.Append .CreateParameter("ssFileDownloadDNSHost", adVarChar, adParamInput, 255, ssFileDownloadDNSHost)
		.Parameters.Append .CreateParameter("ssFileDownloadedFileName", adVarChar, adParamInput, 255, ssFileDownloadedFileName)
		.Parameters.Append .CreateParameter("ssFileDownloadedVersion", adVarChar, adParamInput, 50, ssFileDownloadedVersion)

		On Error Resume Next
		.Execute ,,128	'adExecuteNoRecords
	End With	'pobjCommand
	Set pobjCommand = Nothing
		
	If Err.number <> 0 Then
		If ssDebug_Download Then
			Response.Write "<fieldset><legend>Error in RecordDownloadInitiated</legend>" _
						& "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" _
						& "pstrSQL: " & pstrSQL & "<hr>" _
						& "ssFileDownloadInitiated: " & .Parameters("ssFileDownloadInitiated").Value & "<br />" _
						& "ssFileDownloadOrderItemID: " & .Parameters("ssFileDownloadOrderItemID").Value & "<br />" _
						& "ssFileDownloadREMOTE_ADDR: " & .Parameters("ssFileDownloadREMOTE_ADDR").Value & "<br />" _
						& "ssFileDownloadREMOTE_HOST: " & .Parameters("ssFileDownloadREMOTE_HOST").Value & "<br />" _
						& "ssFileDownloadDNSHost: " & .Parameters("ssFileDownloadDNSHost").Value & "<br />" _
						& "ssFileDownloadedFileName: " & .Parameters("ssFileDownloadedFileName").Value & "<br />" _
						& "ssFileDownloadedVersion: " & .Parameters("ssFileDownloadedVersion").Value & "<br />" _
						& "</fieldset>"
		End If
		Err.Clear
	Else
		pstrSQL = "Select ssFileDownloadID From ssFileDownloads Where ssFileDownloadInitiated=? And ssFileDownloadOrderItemID=? And ssFileDownloadREMOTE_ADDR=?"
		Set pobjCommand = CreateObject("ADODB.Command")
		With pobjCommand
			.CommandType = adCmdText
			.ActiveConnection = cnn
			.CommandText = pstrSQL
			.Parameters.Append .CreateParameter("ssFileDownloadInitiated", adDBTimeStamp, adParamInput, 16, pTempssFileDownloadInitiated)
			.Parameters.Append .CreateParameter("ssFileDownloadOrderItemID", adInteger, adParamInput, 4, ssFileDownloadOrderItemID)
			.Parameters.Append .CreateParameter("ssFileDownloadREMOTE_ADDR", adVarChar, adParamInput, 255, ssFileDownloadREMOTE_ADDR)
			
			Set pobjRS = .Execute
			If Err.number = 0 Then
				If Not pobjRS.EOF Then plngID = pobjRS.Fields("ssFileDownloadID").Value
			Else
				If ssDebug_Download Then
					Response.Write "<fieldset><legend>Error in RecordDownloadInitiated</legend>" _
								& "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" _
								& "pstrSQL: " & pstrSQL & "<hr>" _
								& "ssFileDownloadInitiated: " & .Parameters("ssFileDownloadInitiated").Value & "<br />" _
								& "ssFileDownloadOrderItemID: " & .Parameters("ssFileDownloadOrderItemID").Value & "<br />" _
								& "ssFileDownloadREMOTE_ADDR: " & .Parameters("ssFileDownloadREMOTE_ADDR").Value & "<br />" _
								& "</fieldset>"
				End If
				Err.Clear
			End If	'Err.number = 0
			
			pobjRS.Close
			Set pobjRS = Nothing
		End With	'pobjCommand
		Set pobjCommand = Nothing
	End If	'Err.number <> 0
		
	
	RecordDownloadInitiated = plngID

End Function	'RecordDownloadInitiated

'**********************************************************************************************************

Function RecordDownloadCompleted(byVal ssFileDownloadID)
							
Dim pblnSuccess
Dim pobjCommand

	pblnSuccess = True

	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.ActiveConnection = cnn
		.CommandText = "Update ssFileDownloads Set ssFileDownloadCompleted=? Where ssFileDownloadID=?"

		.Parameters.Append .CreateParameter("ssFileDownloadCompleted", adDBTimeStamp, adParamInput, 16, Now())
		.Parameters.Append .CreateParameter("ssFileDownloadID", adInteger, adParamInput, 4, ssFileDownloadID)

		'On Error Resume Next
		.Execute ,,128	'adExecuteNoRecords
		
		If Err.number <> 0 Then
			If ssDebug_Download Then Response.Write "<font color=red>Error in RecordDownloadCompleted: Error " & Err.number & ": " & Err.Description & "</font><br />" & vbcrlf
			pblnSuccess = False
			Err.Clear
		End If	'Err.number <> 0
		
	End With	'pobjCommand
	Set pobjCommand = Nothing
	
	RecordDownloadCompleted = pblnSuccess

End Function	'RecordDownloadCompleted

'**********************************************************************************************************

Sub updateDownloadLimits(byVal lngOrderDetailID)
'Purpose: Update download limits in place at time of first download
'Assumption: Max download count is null in order details until time limits imposed
'Notes: Placed here to avoid slowing down order confirmation process

Dim pobjCommand
Dim pobjRS

	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.ActiveConnection = cnn
		.CommandText = "SELECT sfProducts.prodMaxDownloads, sfProducts.prodDownloadValidFor" _
					 & " FROM sfOrderDetails INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
					 & " WHERE sfOrderDetails.odrdtMaxDownloads is Null And sfOrderDetails.odrdtID=?"

		.Parameters.Append .CreateParameter("odrdtID", adInteger, adParamInput, 4, lngOrderDetailID)

		'On Error Resume Next
		Set pobjRS = .Execute
		If Err.number = 0 Then
			If Not pobjRS.EOF Then
				.Parameters.Delete("odrdtID")
			
				.CommandText = "Update sfOrderDetails Set odrdtMaxDownloads=?, odrdtDownloadExpiresOn=? WHERE odrdtID=?"
			
				If isNull(pobjRS.Fields("prodMaxDownloads").Value) Then
					.Parameters.Append .CreateParameter("odrdtMaxDownloads", adInteger, adParamInput, 4, 0)
				Else
					.Parameters.Append .CreateParameter("odrdtMaxDownloads", adInteger, adParamInput, 4, pobjRS.Fields("prodMaxDownloads").Value)
				End If
				
				If isNull(pobjRS.Fields("prodDownloadValidFor").Value) Then
					.Parameters.Append .CreateParameter("odrdtDownloadExpiresOn", adDBTimeStamp, adParamInput, 4, Null)
				Else
					.Parameters.Append .CreateParameter("odrdtDownloadExpiresOn", adDBTimeStamp, adParamInput, 4, DateAdd("d", pobjRS.Fields("prodDownloadValidFor").Value, Now()))
				End If
				
				.Parameters.Append .CreateParameter("odrdtID", adInteger, adParamInput, 4, lngOrderDetailID)
				.Execute ,,128	'adExecuteNoRecords
			End If
			pobjRS.Close
		Else
			If ssDebug_Download Then
				Response.Write "<fieldset><legend>Error in updateDownloadLimits</legend>" _
							& "<font color=red>Error " & Err.number & ": " & Err.Description & "</font><br />" _
							& "pstrSQL: " & .CommandText & "<hr>" _
							& "odrdtID: " & .Parameters("odrdtID").Value & "<br />" _
							& "</fieldset>"
			End If
			Err.Clear
		End If	'Err.number = 0
		Set pobjRS = Nothing
		
	End With	'pobjCommand
	Set pobjCommand = Nothing
	
End Sub	'updateDownloadLimits

'**********************************************************************************************************

Sub authorizeDownloadsByOrder(byVal lngOrderID)

Dim pobjCommand

	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.ActiveConnection = cnn
		.CommandText = "Update sfOrderDetails Set odrdtDownloadAuthorized=1 WHERE odrdtOrderId=?"

		.Parameters.Append .CreateParameter("odrdtOrderId", adInteger, adParamInput, 4, lngOrderID)
		.Execute ,,128	'adExecuteNoRecords
	End With	'pobjCommand
	Set pobjCommand = Nothing
	
End Sub	'authorizeDownloadsByOrder

'**********************************************************************************************************

Function orderContainsDownloadableItems(byVal lngOrderID)

Dim pobjCommand
Dim pobjRS

	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.ActiveConnection = cnn
		.CommandText = "SELECT sfOrderDetails.odrdtID, sfProducts.prodFileName, sfOrderDetails.odrdtOrderId" _
					 & " FROM sfOrderDetails INNER JOIN sfProducts ON sfOrderDetails.odrdtProductID = sfProducts.prodID" _
					 & " WHERE (sfProducts.prodFileName Is Not Null) AND (sfProducts.prodFileName<>'') AND (sfOrderDetails.odrdtOrderId=?)"

		.Parameters.Append .CreateParameter("odrdtOrderId", adInteger, adParamInput, 4, lngOrderID)
		Set pobjRS = .Execute
		orderContainsDownloadableItems = Not pobjRS.EOF
		pobjRS.Close
		Set pobjRS = Nothing
	End With	'pobjCommand
	Set pobjCommand = Nothing
	
End Function	'orderContainsDownloadableItems

'**********************************************************************************************************

Function updateDownload_OrderDetail(byVal lngOrderDetailID)
'Purpose: Updates download setting in Order Manager

Dim pobjCommand
Dim pobjRS
Dim odrdtDownloadAuthorized
Dim odrdtMaxDownloads
Dim odrdtDownloadExpiresOn

	odrdtDownloadAuthorized = Trim(Request.Form("odrdtDownloadAuthorized." & Trim(lngOrderDetailID)))
	odrdtMaxDownloads = Trim(Request.Form("odrdtMaxDownloads." & Trim(lngOrderDetailID)))
	odrdtDownloadExpiresOn = Trim(Request.Form("odrdtDownloadExpiresOn." & Trim(lngOrderDetailID)))
	
	If Len(odrdtMaxDownloads) = 0 Or Not isNumeric(odrdtMaxDownloads) Then odrdtMaxDownloads = 0
	If odrdtDownloadAuthorized <> "1" Then odrdtDownloadAuthorized = 0
	If Not isDate(odrdtDownloadExpiresOn) Then odrdtDownloadExpiresOn = ""
	
	If ssDebug_Download Then
		Response.Write "<fieldset><legend>updateDownload_OrderDetail</legend>" _
					& "odrdtDownloadAuthorized: " & odrdtDownloadAuthorized & "<br />" _
					& "odrdtMaxDownloads: " & odrdtMaxDownloads & "<br />" _
					& "odrdtDownloadExpiresOn: " & odrdtDownloadExpiresOn & "<br />" _
					& "odrdtID: " & lngOrderDetailID & "<br />" _
					& "</fieldset>"
	End If

	Set pobjCommand = CreateObject("ADODB.Command")
	With pobjCommand
		.CommandType = adCmdText
		.ActiveConnection = cnn

		.CommandText = "Update sfOrderDetails Set odrdtDownloadAuthorized=?, odrdtMaxDownloads=?, odrdtDownloadExpiresOn=? WHERE odrdtID=?"
	
		.Parameters.Append .CreateParameter("odrdtDownloadAuthorized", adInteger, adParamInput, 4, odrdtDownloadAuthorized)
		.Parameters.Append .CreateParameter("odrdtMaxDownloads", adInteger, adParamInput, 4, odrdtMaxDownloads)
		
		If Len(odrdtDownloadExpiresOn) = 0 Then
			.Parameters.Append .CreateParameter("odrdtDownloadExpiresOn", adDBTimeStamp, adParamInput, 4, Null)
		Else
			.Parameters.Append .CreateParameter("odrdtDownloadExpiresOn", adDBTimeStamp, adParamInput, 4, odrdtDownloadExpiresOn)
		End If
		
		.Parameters.Append .CreateParameter("odrdtID", adInteger, adParamInput, 4, lngOrderDetailID)
		.Execute ,,128	'adExecuteNoRecords
		
	End With	'pobjCommand
	Set pobjCommand = Nothing
	
End Function	'updateDownload_OrderDetail

%>
