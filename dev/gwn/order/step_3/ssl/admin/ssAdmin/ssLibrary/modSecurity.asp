<%
'********************************************************************************
'*   Webstore Manager Version SF 5.0                                            *
'*   Release Version   1.0                                                      *
'*   Release Date      April 13, 2001			                                *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'To use integrated security remove the apostrophe befor the "If"

If Len(Session("login")) = 0 And cblnUseIntegratedSecurity Then Response.Redirect "Admin.asp?PrevPage=" & Request.ServerVariables("SCRIPT_NAME")
'<!--#include file="../../../SFLib/sfsecurity.asp"-->
%>
