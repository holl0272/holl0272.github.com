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

'--------------------------------------------------------------------------------------------------
%>
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/incCC.asp"-->
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_class.asp"-->
<!--#include file="OrderManager_Support/ssOrderAdmin_OrdersToXML.asp"-->
<!--#include file="ssOrderAdmin_common.asp"-->
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

Const filename = "Export.csv"

    maryTemp = Split(LoadRequestValue("Action"), "|")
    
	mstrAction = maryTemp(0)
	If UBound(maryTemp) > 0 Then mstrXSLFilePath = Trim(maryTemp(1))
	If UBound(maryTemp) > 1 Then mstrExportField = Trim(maryTemp(2))
	
    Select Case mstrAction
        Case "downloadOrders"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrXSLFilePath)
			Response.ContentType = "application/octet-stream"
			Response.AddHeader "Content-Disposition", "attachment; filename=""" & filename & """"
			Response.Write exportOrders(Request.Form("chkssOrderID"), mstrXSLFilePath)
        Case "exportOrdersCustom"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrXSLFilePath)
			Response.Write exportOrders(Request.Form("chkssOrderID"), mstrXSLFilePath)
        Case "printOrders"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrXSLFilePath)
			Call exportOrders(Request.Form("chkssOrderID"), mstrXSLFilePath)
			Response.Write "<OBJECT ID=WebBrowser1 WIDTH=0 HEIGHT=0 CLASSID='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2'></OBJECT>"
			Response.Write "<script language=javascript>" _
						   & "document.all('WebBrowser1').ExecWB(6, 2);window.close();</script>"
        Case "viewOrders"
			mstrXSLFilePath = Server.MapPath("exportTemplates/" & mstrXSLFilePath)
			Response.Write exportOrders(Request.Form("chkssOrderID"), mstrXSLFilePath)
    End Select
    
    If Len(mstrExportField) > 0 Then Call setExportedStatus(Request.Form("chkssOrderID"), mstrExportField, 1)
    
    Call ReleaseObject(cnn)
   
%>