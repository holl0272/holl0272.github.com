<%@ LANGUAGE="VBScript" %>

<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront 99 v.3.0.1
'
'   Author     :   LaGarde, Incorporated
'
'   Notes      :  There are no configurable elements in this file.
'                  
'
'                         COPYRIGHT NOTICE
'
'   The contents of this file is protected under the United States
'   copyright laws as an unpublished work, and is confidential and
'   proprietary to LaGarde, Incorporated.  Its use or disclosure in 
'   whole or in part without the expressed written permission of 
'   LaGarde, Incorporated is expressely prohibited.
'
'   (c) Copyright 1998,1999 by LaGarde, Incorporated.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>

<%	
	If Request.ServerVariables("HTTP_COOKIE") = "" Then

        Response.Redirect "error.htm"

	End If


	DSN_Name = Session("DSN_Name")

	Set Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""

	If Request("ORDER_FLAG") = "1" Then

	SQLStmt = "SELECT GRAND_TOTAL from customer WHERE "
	SQLStmt = SQLStmt & " CUSTOMER_ID = " & Session("ORDER_ID") & ""
	RSOrderCheck = Connection.Execute(SQLStmt)
	If RSOrderCheck("GRAND_TOTAL") < "1" Then
	
	
	If Request("Quantity") = "" OR Request("Quantity") is nothing Then
	Quantity = Server.URLEncode("<big>There was no quantity specified for your 	order<br>Please enter a quantity for the desired item.</big>")
	Description = Server.URLEncode("&nbsp;")
	Message = Server.URLEncode("&nbsp;")
	Alert = 1
	ElseIf InStr(Request("Quantity"),"-") > 0 OR InStr(Request("Quantity"),".") > 0 Then
	Quantity = Server.URLEncode("<big>Please enter only whole, ")
	Quantity = Quantity & Server.URLEncode("positive numbers for the<br> ")
	Quantity = Quantity & Server.URLEncode("order quantity</big>")
	Description = Server.URLEncode("&nbsp;")
	Message = Server.URLEncode("&nbsp;")
	Alert = 1
	Else

	
	For Each element In Request.Form
	If InStr(element,"PRODUCT_ID") Then 
	PROD_NAME = element
	End If
	Next

	PRODUCT_ID = Request(""&PROD_NAME&"")	
	
	SQLStmt = " SELECT PRICE, DESCRIPTION, MESSAGE FROM PRODUCT WHERE "
	SQLStmt = SQLStmt & " PRODUCT_ID = '" & PRODUCT_ID & "' "
	Set RSOrder = Connection.Execute(SQLStmt)


	IF RSOrder.BOF OR RSOrder.EOF Then
	
	SQLStmt = "SELECT PRIMARY_EMAIL from admin"
	Set RSAdmin = Connection.Execute(SQLStmt)
	MailTo = "mailto:"&RSAdmin("PRIMARY_EMAIL")

	Quantity = Server.URLEncode("<big>This item was not currently available from")
	Quantity = Quantity & Server.URLEncode("inventory.<br>Please contact the ")
	Quantity = Quantity & Server.URLEncode("<a href="&mailto&">merchant </a>for ")
	Quantity = Quantity & Server.URLEncode("further information.</big>")
	Description = Server.URLEncode("&nbsp;")
	Message = Server.URLEncode("&nbsp;")
	Alert = 1
	Else


	MyCurrSymbol = FormatCurrency(1)
	MyCurrSymbol = Replace((MyCurrSymbol),"1","")
	MyCurrSymbol = Replace((MyCurrSymbol),",","")
	MyCurrSymbol = Replace((MyCurrSymbol),".","")
	MyCurrSymbol = Replace((MyCurrSymbol),"0","")

	ProdPrice = Replace(RSOrder("Price"),MyCurrSymbol,"")	

	ExtPrice = ((ProdPrice)*Request("Quantity"))

	MyCurrFormat = FormatCurrency(1)
	If InStr(MyCurrFormat,",") Then
	rPrice = Replace((ProdPrice),",",".")
	RExtPrice = Replace((ExtPrice),",",".")
	Else 
	rPrice = ProdPrice
	RExtPrice = ExtPrice
	End If
	
	

		
	
	SQLStmt = " INSERT INTO ORDERS(ORDER_ID, PRODUCT_ID,"
	SQLStmt = SQLStmt & " QUANTITY, PRICE, TOTAL) "
	SQLStmt = SQLStmt & " VALUES(" & Session("ORDER_ID") & ","
	SQLStmt = SQLStmt & " '" & PRODUCT_ID & "',"
	SQLStmt = SQLStmt & " " & Request("Quantity") & ","
	SQLStmt = SQLStmt & " " & rPrice & ","
	SQLStmt = SQLStmt & " " & RExtPrice & ")"

	Set RSAddProd = Connection.Execute(SQLStmt)
	
	Quantity = Request("Quantity")
	
	Alert = 0
	
	If IsNull(RSOrder("Description")) Then
	Description = ""
	Else
	Description = Server.URLEncode(RSOrder("Description"))
	End If
	
	If IsNull(RSOrder("Message")) Then
	Message = ""
	Else
	Message = Server.URLEncode(RSOrder("Message"))
	End If
	
	End If
	
	Order_Flag = 1
	End If

	SndPage = Request.ServerVariables("HTTP_REFERER")
	If InStr(SndPage, "?")>0 Then
	SndPage = Left(SndPage, InStr(SndPage, "?") - 1)
	End If
	

	ChkPath = "http://"&Request.ServerVariables("SERVER_NAME")&"/"
	'Response.Write ChkPath
	'Response.Write SndPage
	If SndPage = ChkPath Then
	SndPage = "http://"&Request.ServerVariables("SERVER_NAME")&"/default.asp"
	End If
	

	SRCH_DESCRIPTION = Server.URLEncode(Request("SRCH_DESCRIPTION"))
	SRCH_MANUFACTURER = Server.URLEncode(Request("SRCH_MANUFACTURER"))
	SRCH_ID = Server.URLEncode(Request("SRCH_ID"))
	SRCH_CATEGORY = Server.URLEncode(Request("SRCH_CATEGORY"))
	RowCount = Server.URLEncode(Request("RowCount"))
	Order_Flag = 1
	'RowCount = RowCount - 1
	PageNo = Request("PageNo")

	ReturnPg = SndPage&"?Order_Flag="&Order_Flag&"&Quantity="&Quantity&"&Description="&Description&"&Message="&Message&"&Alert="&Alert&"&SRCH_DESCRIPTION="&SRCH_DESCRIPTION&"&SRCH_MANUFACTURER="&SRCH_MANUFACTURER&"&SRCH_ID="&SRCH_ID&"&SRCH_CATEGORY="&SRCH_CATEGORY&"&PageNo="&PageNo&"&RowCount="&RowCount
	Connection.Close
	Response.Redirect ReturnPg 
	
	Else 
	Response.Redirect "order_complete.asp?DSN_NAME="&DSN_NAME
	Connection.Close
	End If
	End If
 %>