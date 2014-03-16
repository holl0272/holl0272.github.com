<%@ LANGUAGE="VBScript" %><%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   System     :   StoreFront v.3.0
'
'   Author     :   LaGarde, Incorporated
'   Description:   Builds the customer's order by adding the selected product
'                  to the orders table.
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
'   (c) Copyright 1998 by LaGarde, Incorporated.  All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
%><%	
	
	'Determines the DSN for the web store.  This is a system variable that
	'is created in the global.asa file.


	DSN_Name = Session("DSN_Name")

	SET Connection = Server.CreateObject("ADODB.Connection")
	
	Connection.Open "DSN="&DSN_Name&""

SQL = "ALTER TABLE admin ADD PAYMENT_SERVER Char(100)"
SET AD_PYMNT_SERVER = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD LOGIN Char(25)"
SET AD_LOGIN = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD MERCHANT_TYPE Char(25)"
SET AD_MERCH_TYPE = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD STATE_TAX_AMOUNT Char(5)"
SET AD_ST_TAX = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD DOMAIN_NAME Char(100)"
SET AD_DOMAIN = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD COUNTRY_TAX_AMOUNT Char(5)"
SET AD_CNT_TAX = Connection.Execute(SQL)

SQL = "ALTER TABLE admin DROP CCPAYMENT_SERVER"
SET DP_PYMT_SER = Connection.Execute(SQL)

SQL = "ALTER TABLE admin DROP CC_SECRET"
SET DP_SECRET = Connection.Execute(SQL)

SQL = "ALTER TABLE admin DROP CC_MERCHANT_TYPE"
SET DP_MERCH_TYPE = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD MAIL_SERVER Char(100)"
SET AD_REFERER = Connection.Execute(SQL)

SQL = "ALTER TABLE admin DROP TAX_AMOUNT"
SET DP_TAX_AMT = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD PAYMENT_METHOD Char(30)"
SET AD_PAY_MTHD = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD BANK_NAME Char(255)"
SET AD_BK_NM = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD ROUTING_NO Char(255)"
SET AD_BK_RT = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD CHK_ACCT_NO Char(20)"
SET AD_CK_ACCT = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD PURCH_ORDER_NO Char(50)"
SET AD_PRCH_ORD = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD REFERER Char(100)"
SET AD_REFERER = Connection.Execute(SQL)

SQL = "ALTER TABLE customer ADD HTTP_REFERER Char(100)"
SET AD_HTTP_REFERER = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD MAIL_CC Char(5)"
SET AD_MAIL_CC = Connection.Execute(SQL)

SQL = "ALTER TABLE admin ADD MAIL_METHOD Char(25)"
SET AD_MAIL_METH = Connection.Execute(SQL)

SQL = "DROP TABLE cybercash"
SET DP_TBL_CC = Connection.Execute(SQL)

SQL = "CREATE TABLE transactions (ID Number, ORDER_ID Number, ORDER_DATE Date, CUST_TRANS_NO Char(25), MERCH_TRANS_NO Char(25), AVS_CODE Char(25), AUX_MSG Char(25), ACTION_CODE Char(25), RETRIEVAL_CODE Char(25), AUTH_NO Char(25), ERROR_MSG Char(25), ERROR_LOCATION Char(25), STATUS Char(25))"
SET CRT_TRANS = Connection.Execute(SQL)

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charSET=windows-1252">
<title>AuthorizeNet Update</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>

<p> Your StoreFront database has been successfully updated.&nbsp; The new features
included in this update are the ability to SET both a country or national tax rate as well
as a state tax rate.&nbsp; This update also provides built-in support for several new
transaction processing services. In addition to support for CyberCash 2.1.4 which was
included in StoreFront Standard and StoreFront Pro versions 1.0 to 2.5, this newest
version of StoreFront includes built-in support for CyberCash 3.2, AuthorizeNet
(www.authorizenet.com) and CarmelCash (wwww.carmelww.com) transaction services.&nbsp;
Support for accepting electronic checks is also built into this newest version of
StoreFront.</p>

<p> To reset your StoreFront web store for the correct use of the new features, please go
into the <a href="menu.asp">new store Setup </a>and enter the new configuration information 
required for support of these new features. </p>
<% Connection.Close %>
</body>
</html>
