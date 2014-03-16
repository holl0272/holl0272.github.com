<%
Option Explicit
Response.Buffer = False
Server.ScriptTimeout = 900

'********************************************************************************
'*   Customer Manager for StoreFront 5.0                                        *
'*   Release Version:	2.00.004		                                        *
'*   Release Date:		August 18, 2003											*
'*   Revision Date:		March 16, 2005											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   2.00.004 (March 16, 2005)													*
'*   - Enhancement - added tabbed interface										*
'*                                                                              *
'*   2.00.003 (January 14, 2004)                                                *
'*   - Bug fix - update routine modified to use nulls instead of empty values   *
'*                                                                              *
'*   2.00.002 (November 6, 2003)                                                *
'*   - Added Pricing Level Manager support                                      *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2001 Sandshot Software.  All rights reserved.                *
'********************************************************************************

'/////////////////////////////////////////////////
'/
'/  User Parameters
'/

	Const clngMaxToUse = 100
	Const clngMinToUse = 3
	Const cstrFileLocation = "..\..\..\fpdb\InStockImport.csv"
	
	Const cstrEmailFrom = "from@sandshot.net"
	Const cstrEmailCC = "cc@sandshot.net"
	Const cstrEmailBCC = "bcc@sandshot.net"
	Const emailTemplateDirectory = "emailTemplates\NotifyMe\"
	Dim maryEmailTemplates:	maryEmailTemplates = Array("NotifyMe1st.txt", "NotifyMe2nd.txt")

'/
'/////////////////////////////////////////////////

'**********************************************************
'	Developer notes
'**********************************************************

'**********************************************************
'*	Functions
'**********************************************************

'Function getInventoryUpdateFile()
'Function LoadProductByInventoryID(byVal lngInventoryID, byRef objRS)

'**********************************************************
'*	Page Level variables
'**********************************************************

Dim i
Dim maryEmails
Dim maryInventory
Dim maryNotifyMe
Dim mclsReplacement
Dim mlngInventory
Dim mlngInventory_Current
Dim mlngInventoryID
Dim mobjRS
Dim mstrAction

'**********************************************************
'*	Begin Page Code
'**********************************************************

Function getInventoryUpdateFile()

Dim paryInventory
Dim f
Dim fs
Dim i
Dim path
Dim str_CurrentLine
Dim paryTemp

	path = Server.MapPath(cstrFileLocation)
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set f=fs.OpenTextFile(path, 1)

	i = -1
	'Response.Write "<ol>"
	Do While Not f.AtEndOfStream
		str_CurrentLine = f.ReadLine
		'Response.Write "<li>" & str_CurrentLine & "</li>"
		If Len(str_CurrentLine) > 0 Then i = i + 1
	Loop
	'Response.Write "</ol>"
	'Response.Flush
	
	ReDim paryInventory(i)
	
	Set f=fs.OpenTextFile(path, 1)

	i = -1
	Do While Not f.AtEndOfStream
		str_CurrentLine = f.ReadLine
		If Len(str_CurrentLine) > 0 Then
			i = i + 1
			paryTemp = Split(str_CurrentLine, ",")
			paryInventory(i) = Array(paryTemp(0), paryTemp(1))
		End If
	Loop

	f.Close
	Set f=Nothing
	Set fs=Nothing
	
	getInventoryUpdateFile = paryInventory
	
End Function	'getInventoryUpdateFile

'**********************************************************

Function LoadProductByInventoryID(byVal lngInventoryID, byRef objRS)

Dim pstrSQL

	If Len(lngInventoryID) > 0 And isNumeric(lngInventoryID) Then
		pstrSQL = "Select invenProdId, invenAttDetailID, invenAttName, invenInStock, invenLowFlag From sfInventory Where invenId=" & lngInventoryID
		Set objRS = GetRS(pstrSQL)
		If objRS.EOF Then
			LoadProductByInventoryID = False
		Else
			LoadProductByInventoryID = True
		End If
	Else
		LoadProductByInventoryID = False
	End If

End Function	'LoadProductByInventoryID

'***********************************************************************************************

Function SendInventoryInStockNotification(byRef aryNotifyMe, byRef objclsReplacement)

Dim i
Dim pclsEmail
Dim plngPriorNotifications
Dim pstrEmailMessage
Dim pstrEmailTo
Dim pstrSubject
Dim pstrFileName

	plngPriorNotifications = aryNotifyMe(10)
	
	If plngPriorNotifications > UBound(maryEmailTemplates) Then
		pstrFileName = maryEmailTemplates(UBound(maryEmailTemplates))
	Else
		pstrFileName = maryEmailTemplates(plngPriorNotifications)
	End If
	
	pstrEmailTo = aryNotifyMe(6)
	'Response.Write pstrEmailTo & " (" & plngPriorNotifications & ")"

	If LoadEmailTemplates(emailTemplateDirectory, maryEmails) Then
		For i = 0 To UBound(maryEmails)
			If CBool(maryEmails(i)(enEmail_FileName) = pstrFileName) Or CBool((i = UBound(maryEmails)) And (Len(pstrFileName) = 0)) Then
				pstrSubject = objclsReplacement.getReplacment(maryEmails(i)(enEmail_Subject))
				pstrEmailMessage = objclsReplacement.getReplacment(maryEmails(i)(enEmail_Body))
				
				Set pclsEmail = New clsEmail
				With pclsEmail
					.MailMethod = adminMailMethod
					.MailServer = adminMailServer
					.ShowFailures = True
					
					If Len(cstrEmailFrom) = 0 Then
						.From = adminPrimaryEmail
					Else
						.From = cstrEmailFrom
					End If
					.To = pstrEmailTo
					.CC = cstrEmailCC
					.BCC = cstrEmailBCC
					.Subject = pstrSubject
					.Body = pstrEmailMessage
					.Send
					
				End With
				Set pclsEmail = Nothing
				SendInventoryInStockNotification = True
				Exit Function
			End If
		Next 'i
	End If

	SendInventoryInStockNotification = False
	
End Function	'SendInventoryInStockNotification

'***********************************************************************************************

Function UpdateNotification(byRef aryNotifyMe)

Dim pobjCmd

	Set pobjCmd  = CreateObject("ADODB.Command")
	With pobjCmd
		.Commandtype = adCmdText
		Set .ActiveConnection = cnn
		.Commandtext = "Update notifyMe Set notifyDateNotified=?, notifyNotifyCount=? Where notifyMeID=?"
		.Parameters.Append .CreateParameter("notifyDateNotified", adDBTimeStamp, adParamInput, 16, Date())
		.Parameters.Append .CreateParameter("notifyNotifyCount", adInteger, adParamInput, 4, aryNotifyMe(10) + 1)
		.Parameters.Append .CreateParameter("notifyMeID", adInteger, adParamInput, 4, aryNotifyMe(0))
		
		.Execute,,128
		If Err.number = 0 Then
			UpdateNotification = True
		Else
			UpdateNotification = False
		End If

	End With	'pobjCmd
	Set pobjCmd = Nothing

End Function	'UpdateNotification

'***********************************************************************************************

Sub closeObj(byRef objItem)

On Error Resume Next

	objItem.Close
	Set objItem = Nothing	
	If Err.number <> 0 Then Err.Clear

End Sub	'closeObj

'***********************************************************************************************

%>
<!--#include file="ssLibrary/modDatabase.asp"-->
<!--#include file="ssLibrary/modSecurity.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<!--#include file="../SFLib/ssclsDebug.asp"-->
<!--#include file="../SFLib/ssmodNotifyMe.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
	mstrAction = LoadRequestValue("Action")
	maryInventory = getInventoryUpdateFile
	Call WriteHeader("",True)
	
	Set mclsReplacement = New clsReplacement

%>
<span class="pagetitle2">Data to import: <%= Replace(cstrFileLocation, "..\", "") %></span>
<form action="" method="post">
<input type="checkbox" name="Action" id="Action" value="Import">&nbsp;<label for="Action">Check to Import</label> <input type="submit" class="butn" name="btnSubmit" value="Submit">
</form>
<ul style="MARGIN-LEFT: 18pt; MARGIN-RIGHT: 0pt; MARGIN-BOTTOM: 0;">
<li><a href="ssInventoryList.asp">List Inventory</a></li>
</ul>
<table class='tbl' style="border-collapse: collapse" cellpadding='2' cellspacing='0' border='1' bgcolor='whitesmoke' id='tblSummary'>
<colgroup>
  <col align="right" />
  <col align="center" />
  <col align="left" />
  <col align="left" />
  <col align="center" />
  <col align="center" />
  <col align="left" />
  <col align="left" />
</colgroup>
<tr class="tblhdr">
<th>#</th>
<th>ID</th>
<th>Product</th>
<th>Attributes</th>
<th>Qty In Stock</th>
<th>New Qty</th>
<th>Result</th>
<th>Notifications</th>
</tr>
<%
For i = 0 To UBound(maryInventory)
	mlngInventory = maryInventory(i)(1)
	mlngInventoryID = maryInventory(i)(0)
	If isNumeric(mlngInventory) Then mlngInventory = CDbl(mlngInventory)

	If LoadProductByInventoryID(mlngInventoryID, mobjRS) Then
		mlngInventory_Current = mobjRS.Fields("invenInStock").Value
	
		Response.Write "<tr>"
		Response.Write "<td>" & i & ".&nbsp;</td>"
		Response.Write "<td>" & mlngInventoryID & "</td>"
		Response.Write "<td>" & mobjRS.Fields("invenProdId").Value & "</td>"
		Response.Write "<td>" & mobjRS.Fields("invenAttName").Value & "&nbsp;</td>"
		Response.Write "<td>" & mlngInventory_Current & "</td>"
		Response.Write "<td>" & mlngInventory & " "
		
		If CDbl(mlngInventory) > clngMaxToUse Then
			Response.Write "<font color=red>Set to Max (" & clngMaxToUse & ")</font>"
			mlngInventory = clngMaxToUse
		End If
		
		If CDbl(mlngInventory) < clngMinToUse Then
			Response.Write "<font color=red>Set to Zero (" & clngMinToUse & ")</font>"
			mlngInventory = 0
		End If
		Response.Write "</td>"
		
		Response.Write "<td>"
		If mstrAction = "Import" Then
			If mlngInventory_Current <> mlngInventory Then
				cnn.Execute "Update sfInventory Set invenInStock=" & mlngInventory & " Where invenId=" & mlngInventoryID,,128
				Response.Write "<img src=""images/shipped.ico"" border=""0"">"
			Else
				Response.Write "-"
			End If
			
			Response.Write "</td>"
			Response.Write "<td>"
			
			On Error Goto 0
			If LoadNotifyMeByInventoryID(mlngInventoryID, maryNotifyMe) Then
				mclsReplacement.setReplacement "{productName}", getProductInfo(maryNotifyMe(2), enProduct_Name)
				
				If LCase(Left(getProductInfo(maryNotifyMe(2), enProduct_Link), 4)) <> "http" Then
					mclsReplacement.setReplacement "{productLink}", Application("adminDomainName") & getProductInfo(maryNotifyMe(2), enProduct_Link)
				Else
					mclsReplacement.setReplacement "{productLink}", getProductInfo(maryNotifyMe(2), enProduct_Link)
				End If
				
				If SendInventoryInStockNotification(maryNotifyMe, mclsReplacement) Then
					Response.Write maryNotifyMe(6)
					'If UpdateNotification(maryNotifyMe) Then Response.Write " <img src=""images/shipped.ico"" border=""0"">"
					If DeleteNotifyMeByID(maryNotifyMe(0)) Then Response.Write " <img src=""images/shipped.ico"" border=""0"">"
				Else
				End If
			Else
				Response.Write "-"
			End If
		Else
			Response.Write "&nbsp;"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
	Else
		Response.Write "<tr>"
		Response.Write "<td>" & i & "</td>"
		Response.Write "<td>" & mlngInventoryID & "</td>"
		Response.Write "<td colspan=6><font color=red>Does not exist</font></td>"
		Response.Write "</tr>"
	End If
Next
%>
</table>

</body>
</html>
<%
	Set mclsReplacement = Nothing
	Call ReleaseObject(mobjRS)
	Call ReleaseObject(cnn)
%>