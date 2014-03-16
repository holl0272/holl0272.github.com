<%
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version   2.0                                                      *
'*   Release Date      July 4, 2002				                                *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

' This file adjusts for variations in your StoreFront paths.
' No changes are necessary for a standard installation.

dim mrs
dim mstrDiffPath
dim mstrBaseHRef, mstrBasePath

'	set mrs = server.CreateObject("ADODB.RECORDSET")
'	Set mrs = GetRS("Select adminDomainName, adminSSLPath from sfAdmin")
'	if not mrs.eof Then 
'		mstrDiffPath = replace(trim(mrs("adminSSLPath").value),trim(mrs("adminDomainName").value),"")
'		mstrDiffPath = replace(mstrDiffPath,"ssl/process_order.asp","")
'		mstrDiffPath = replace(mstrDiffPath,"/","",1,1)
'	end if
'	mrs.Close
'	set mrs = Nothing

	mstrDiffPath = ""

	If len(mstrDiffPath) = 0 Then
		mstrBaseHRef = "http://" & Request.ServerVariables("SERVER_NAME") & "/"
		mstrBasePath = Request.ServerVariables("APPL_PHYSICAL_PATH")
	Else
		mstrBaseHRef = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & mstrDiffPath & "/"
		mstrBasePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & mstrDiffPath & "\"
	End If
%>
