<%
'********************************************************************************
'*   Order Manager							                                    *
'*   Release Version:   2.00.0001												*
'*   Release Date:		November 15, 2003										*
'*   Release Date:		November 15, 2003										*
'*                                                                              *
'*   Release Notes:                                                             *
'*	 ' Initial Release
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2003 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	
'/
'/////////////////////////////////////////////////

Dim cTrackingColumn_OrderNumber
Dim cTrackingColumn_ShippingMethod
Dim cTrackingColumn_TrackingNumber
Dim cTrackingColumn_Message
Dim cTrackingColumn_ShipmentDate
Dim cTrackingColumn_SendEmail
	
Dim cblnDefault_SendEmail
Dim cstrDefault_ImportProfile
Dim cstrDefaultOrderShippedEmailTemplate
Dim cstrDefault_CarrierID
Dim cTrackingOrderNumberPrefix
'Dim maryOrderStatuses
Dim maryImportTemplates
Dim cstrEmailTemplateFolder

Call LoadOrderManagerConfigurationSettings

'***********************************************************************************************

Sub setPredefinedImportTemplates(byVal strTemplate)

Dim i

	For i = 1 To UBound(maryImportTemplates)
		If maryImportTemplates(i)(0) = strTemplate Then Exit For
	Next 'i
	If i > UBound(maryImportTemplates) Then i = 0
	
	cstrDefault_CarrierID = maryImportTemplates(i)(1)
	cTrackingColumn_OrderNumber = maryImportTemplates(i)(2)
	cTrackingColumn_TrackingNumber = maryImportTemplates(i)(3)
	cTrackingColumn_ShipmentDate = maryImportTemplates(i)(4)
	cTrackingColumn_ShippingMethod = maryImportTemplates(i)(5)
	cTrackingColumn_Message = maryImportTemplates(i)(6)
	cTrackingColumn_SendEmail = maryImportTemplates(i)(7)

End Sub	'setPredefinedImportTemplates

'***********************************************************************************************

Function downloadFileName(byVal strTemplateName)

Dim pstrFileName
Dim pdtDate
Dim pstrDay
Dim pstrMonth

	Select Case LCase(strTemplateName)
		Case "dotcomdistribution.xsl"
			pdtDate = Date()
			pstrDay = Mid(CLng(Day(pdtDate) + 100), 2)
			pstrMonth = Mid(CLng(Month(pdtDate) + 100), 2)
			
			pstrFileName = "clientID" & "ord" & pstrMonth & pstrDay & "1"
			
		Case Else:
			pstrFileName = cstrDefaultExportFilename
	End Select
	
	downloadFileName = pstrFileName
    
End Function	'downloadFileName

'***********************************************************************************************

Sub LoadOrderManagerConfigurationSettings

	cstrDefaultOrderShippedEmailTemplate = getAddonConfigurationSetting("OrderManager", "cstrDefaultOrderShippedEmailTemplate")
	cstrEmailTemplateFolder = getAddonConfigurationSetting("OrderManager", "cstrEmailTemplateFolder")
	cstrDefault_ImportProfile = getAddonConfigurationSetting("OrderManager", "cstrDefault_ImportProfile")
	cblnDefault_SendEmail = ConvertToBoolean(getAddonConfigurationSetting("OrderManager", "cblnDefault_SendEmail"), False)
	

	cstrDefault_CarrierID = getAddonConfigurationSetting("OrderManager", "cstrDefault_CarrierID")
	cTrackingOrderNumberPrefix = getAddonConfigurationSetting("OrderManager", "cTrackingOrderNumberPrefix")

	'maryOrderStatuses = getAddonConfigurationSetting_OrderStatuses
	maryImportTemplates = getAddonConfigurationSetting_ImportProfiles
	
End Sub	'LoadOrderManagerConfigurationSettings
%>