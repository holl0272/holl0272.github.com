<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="incCoreFiles.asp"-->
<!--#include file ="ssl/SFLib/incLogin.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.2

'@FILENAME: login.asp
 

'

'@DESCRIPTION: Handles customer login

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
%>
<%
Dim sFirstTime, sReturn, sFormActionType, sThanks, sOutput, sChangeUser, sCustomerForm, sSubmitAction
Dim iCookieID, sForgotPwd, sSendEmail, iAuthenticate, bSvdCartCustomer, bEmailAddress, sNewAccount, sChangeCart, sChange, sDirection, sSvdCartCustEmail
Dim sLoggedIn, sEmail, sPassword, sLogin

'------------------------------------------------------------
' For savecart
'------------------------------------------------------------

' Check which button has been depressed
sFirstTime	  	= Trim(Request.Form("SignUp.x"))
sLoggedIn	  	= getCookie_SessionID
sReturn	  	= Trim(Request.Form("Return.x"))
sChangeUser 	= Trim(Request.Form("ChangeUser.x"))
sForgotPwd	  	= Trim(Request.QueryString("FPWD")) ' Request.Form("FPWD.x")
sSendEmail	  	= Trim(Request.Form("SendEmail.x"))
sNewAccount  	= Trim(Request.QueryString("New")) '.Form("New.x")
If Len(sNewAccount) = 0 Then sNewAccount = Trim(Request.QueryString("Type"))
sChangeCart = Trim(Request.Form("ChangeCart.x"))
sChange		= Trim(Request.Form("Change.x"))
sLogin = Request.Cookies("sfThanks")("PreviousAction")
sEmail	  = Trim(Request.Form("Email"))
sPassword = Trim(Request.Form("Passwd"))
iCustID = custID_cookie

' For people already logged in 
If (CStr(sLoggedIn) = CStr(SessionID) AND Len(CStr(iCustID)) > 0 And sChangeCart = "" And sChange = "" AND sForgotPwd = "") Then
	
	' Possibility - get a new account 
	If sNewAccount <> "" Then	   
		sFormActionType = "NewAccount"
	
	' Through new account, get them a new account   
	ElseIf sFirstTime <> "" Then	
		sFormActionType = "FirstTime"
	
	' If there is email and password, authenticate
	ElseIf Len(sEmail) > 0 And Len(sPassword) > 0 Then
	
		iAuthenticate = customerAuth(sEmail, sPassword, "loose")
		If iAuthenticate > 0 Then
			' Check if there is a custID
			If custID_cookie = "" Then
				sFormActionType = "Returning"	
			Else	
				If custID_cookie <> iAuthenticate Then
					Call setCookie_custID(iAuthenticate, Date() + 730)
				End If		
				' Redirect to proc_order
				Call cleanup_dbconnopen	'This line needs to be included to close database connection
				Response.Redirect(C_SecurePath)
			' End custId if		
			End If 		  
		Else
			sFormActionType = "FirstTime"
			' End Authenticate If
		End If		 
		   
	Else
		If custID_cookie <> "" Then
			If sLogin = "FromShopCart" Then	
				sFormActionType = "NewAccount"
			Else
				Call cleanup_dbconnopen	'This line needs to be included to close database connection
				Response.Redirect(C_SecurePath)
			End If
		Else
			sFormActionType	= "Returning"	
		End If		 	
			
	' End auth if		
	End If			
	
Else 
	If sFirstTime <> "" Then
		sFormActionType = "FirstTime"
	ElseIf sReturn <> "" Then
		sFormActionType = "Returning"	
	ElseIf sForgotPwd <> "" Then
		sFormActionType = "ForgotPwd"
	ElseIf sSendEmail <> "" Then
		sFormActionType = "SendEmail"	
	ElseIf sNewAccount <> "" Then
		sFormActionType = "NewAccount"	
	ElseIf sChangeCart <> "" Then
		sFormActionType = "ChangeCart"	
	ElseIf sChange <> "" Then
		sFormActionType = "Change"	
	End If
End If

'-----------------------------------------------------------
' Cases for saved cart
'-----------------------------------------------------------	
	
	If vDebug = 1 Then Response.Write "<br />FormActionType: " & sFormActionType	
		
	' First time at login, no actions	
	If sFormActionType = "" And custID_cookie = "" Then		
		sOutput = "General"
		
	ElseIf sFormActionType = "" And custID_cookie <> "" Then		
		sOutput = "HasID"
 
	ElseIf sFormActionType = "NewAccount" Then		
		sOutput = "NewAccount"
	
	' Forgot password
	ElseIf sFormActionType = "ForgotPwd" Then		
		sOutput = "Email"		
	
	ElseIf sFormActionType = "ChangeCart" Then
		sOutput = "ChangeCart"
	
	ElseIf sFormActionType = "Change" Then
		iAuthenticate = customerAuth(sEmail,sPassword,"loose")
		If iAuthenticate > 0 Then
			Call cleanup_dbconnopen	'This line needs to be included to close database connection
			Response.Redirect("savecart.asp")
		Else 
			sOutput = "FailedAuthChange"
		End If		
	
	' Send Mail
	ElseIf sFormActionType = "SendEmail" Then		
		sEmail = Request.Form("Email")
		' Send email with password, returns a success or failure boolean
		bEmailAddress = SendPassword(sEmail)
		
		If bEmailAddress = 1 Then
			sOutPut = "SentEmail"
		Else
			sOutPut = "FailedEmail"
		End If		
		
	' First time login		
	ElseIf sFormActionType = "FirstTime" Then	
		
		' Check if email and password correspond to any customer on record	
			iCustID = customerAuth(sEmail, sPassword, "loose")	
			If iCustID > 0 Then
				' For Customers from SaveCart, special case
				If sLogin <> "" Then
		
					' Update Saved Table with iCustID in Table
					Call UpdateCustID(iCustID)							
					If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then
						' Redirect to thanks page
						Response.Cookies("sfThanks").Expires = NOW()
						'Response.Redirect "addproduct.asp?logedin=1"
						Call cleanup_dbconnopen	'This line needs to be included to close database connection
						Response.Redirect getLastSearch
					ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
						' Delete from temp										
						Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
						Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
						Response.Cookies("sfThanks").Expires = NOW()
						Call cleanup_dbconnopen	'This line needs to be included to close database connection
						Response.Redirect "order.asp" 					
					End If	
					
				' End SaveCart New customers If
				End If 	

			ElseIf customerAuth(sEmail,sPassword,"loosest") > 0 Then			
				' Email match. Prompt for new email				  
				sOutput = "EmailMatch"
				iCustID = ""	'rest
			Else		
				
				' Write to customer table, write a cookie, and get back the id							
				iCustID = getCustomerID(sEmail,sPassword)	
				If vDebug = 1 Then Response.write "CustID:" & iCustID
			   
				Call SetSessionLoginParameters(iCustID, sEmail)
	
				' For Customers from SaveCart, special case
				If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
		
					' Update Saved Table with iCustID in Table
					Call UpdateCustID(iCustID)							
					
					If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then							
						Response.Cookies("sfThanks").Expires = NOW()
						' Redirect to thanks page
						'Response.Redirect "addproduct.asp?logedin=1"
						Call cleanup_dbconnopen	'This line needs to be included to close database connection
						Response.Redirect getLastSearch
					ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
						' Delete from temp
						Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
						Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
						Response.Cookies("sfThanks").Expires = NOW()
						Call cleanup_dbconnopen	'This line needs to be included to close database connection
						Response.Redirect "order.asp"								
					End If								
					
				' End SaveCart New customers If
				End If 	
		
				' Assume new customer checkout for other cases
				' Redirect to form to enter customer Info
				Call cleanup_dbconnopen	'This line needs to be included to close database connection
				Response.Redirect("order.asp")
					
		' End existing cookie if			
		  End If			

	ElseIf sFormActionType = "Returning" Then	
		sEmail = Trim(Request.Form("Email"))
		sPassword = Trim(Request.Form("Passwd"))
		
		' Authenticate customer
		iAuthenticate = customerAuth(sEmail,sPassword,"loose")
			
		If iAuthenticate <> "" AND iAuthenticate > 0 Then
		
			' Associate sessionID with cookie
			Call setCookie_SessionID(SessionID, Date() + 1)
			
			' Check if customer still has a cookie				
			iCustID = custID_cookie				
		
			' Write to cookie if none exists for custID	
			If iCustID <> iAuthenticate Or Len(iCustID) = 0 Then
				Call setCookie_custID(iAuthenticate, Date() + 730)
				iCustID = iAuthenticate
			End If			
				
			' For Customers from SaveCart, special case
			If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
							
				' Update Saved Table with iCustID in Table
				Call UpdateCustID(iCustID)

				If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then
					' Redirect to thanks page
					Response.Cookies("sfThanks").Expires = NOW()
					'Response.Redirect "addproduct.asp?logedin=1" 
					Call cleanup_dbconnopen	'This line needs to be included to close database connection
					Response.Redirect getLastSearch
				ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
					' Delete from temp
					Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
					Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
					Response.Cookies("sfThanks").Expires = NOW()
					Call cleanup_dbconnopen	'This line needs to be included to close database connection
					Response.Redirect "order.asp"								
				End If	
				
			' End SaveCart New customers If
			End If 
			
			' Check if it is a savedcart customer
			bSvdCartCustomer = getSvdCartCustomer(iCustID,"boolean")
		
			If bSvdCartCustomer = 1 Then
				If vDebug = 1 Then Response.Write "<br /> Wish List Customer " & iCustID
				Call cleanup_dbconnopen	'This line needs to be included to close database connection
				Response.Redirect("order.asp")					
			Else
				' Redirect to proc_order
				Call cleanup_dbconnopen	'This line needs to be included to close database connection
				Response.Redirect("savecart.asp")
			End If 
				
		Else
			' Assume new person
			' Write to customer table, write a cookie, and get back the id							
			iCustID = getCustomerID(sEmail,sPassword)	
			If vDebug = 1 Then Response.write "CustID:" & iCustID
			   
			Call setCookie_SessionID(SessionID, Date() + 1)

			' For Customers from SaveCart, special case
			If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
	
				' Update Saved Table with iCustID in Table
				Call UpdateCustID(iCustID)							
				
				If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then							
					Response.Cookies("sfThanks").Expires = NOW()
					' Redirect to thanks page
					'Response.Redirect "addproduct.asp?logedin=1" 
					Call cleanup_dbconnopen	'This line needs to be included to close database connection
					Response.Redirect getLastSearch
				ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
					' Delete from temp
					Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
					Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
					Response.Cookies("sfThanks").Expires = NOW()
					Call cleanup_dbconnopen	'This line needs to be included to close database connection
					Response.Redirect "order.asp"								
				End If								
				
			' End SaveCart New customers If
			End If 	
				
			' If all else fails, just go to order.asp
			Call cleanup_dbconnopen	'This line needs to be included to close database connection
			Response.Redirect("order.asp")  
		End If	

	' End FormAction If	
	End If	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><%= C_STORENAME %> Login Page</title>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Pragma" content="no-cache">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="keywords">
<meta name="description" content="description">
<meta name="Robot" content="Index,ALL">
<meta name="revisit-after" content="15 Days">
<meta name="Rating" content="General">
<meta name="Language" content="en">
<meta name="distribution" content="Global">
<meta name="Classification" content="classification">

<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">

<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js" type="text/javascript"></script>
</head>
<body <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <%
    '-----------------------------------------------------------
    ' Begin OutPut Block
    '-----------------------------------------------------------
    %>    
     
        <tr>
          <td   class="tdMiddleTopBanner">
  	        <%If sOutPut = "General" Then %>  
			Please Login
	        <%ElseIf sOutPut = "NewAccount" Then %>
			New Account 		
	        <%ElseIf sOutPut = "EmailMatch" Then %>	
			Matching Email Found			
	        <%ElseIf sOutPut = "HasID" Then %>  
			Returning Customer Login		
	        <%ElseIf sOutPut = "FailedAuth" Then %>	
			Failed Authentication
	        <%ElseIf sOutPut = "FailedAuthChange" Then %>	
			Change Cart Failed Authentication			
	        <%ElseIf sOutPut = "Email" Then %>	
			Send Password to Email Address
	        <%ElseIf sOutPut = "SentEmail" Then %>
			Email Sent to Address		
	        <%ElseIf sOutPut = "FailedEmail" Then %>
			No Email Found On Record 				
	        <%ElseIf sOutPut = "ChangeCart" and cblnSF5AE Then %>	
			Change Wish List	
			<%ElseIf sOutPut = "ChangeCart" and Not cblnSF5AE Then %>	
			Change Saved Cart	
	        <%Else%>
	        Login with Email and Password
	        <%
	End If 
	
	If sOutPut = "FailedAuth" Then
		sSubmitAction = "" 
	Else
		sSubmitAction = "return sfCheck(this)"
	End If 
	%>	
		  </td>
		</tr>
		
		<% If sOutPut <> "Email" AND sOutPut <> "SentEmail"  AND sOutPut <> "FailedEmail" Then %>
		<tr>
		  <td class="tdBottomTopBanner">
		    <%If sOutPut = "NewAccount" Then %>  
			 Your Saved Order is your personal, private collection of items that you are thinking about purchasing. Put as many items in your Saved Order as you want. When you decide to buy them, simply add them to your Current Order.<P>To access an existing Saved Order, log in with the e-mail and password you chose when you first signed in. If you've forgotten your password, click on
            Forgot Password for help. To create a new Saved Order, select New Account. 
            <% Elseif sOutPut = "General" Then %>
            Your Saved Order is your personal, private collection of items that you are thinking about purchasing. Put as many items in your Saved Order as you want. When you decide to buy them, simply add them to your Current Order.
           	 To create a new Saved Order, select New Account.            
		    <%ElseIf sOutPut = "HasID" Then %>  
            Please login with the e-mail address and password you chose when you first signed in. If you've forgotten your password, click on
            Forgot Password for help
		    <%ElseIf sOutPut = "FailedAuth" Then %>		
			We're sorry, but your combination of e-mail address and password is not recognized. If you've forgotten your password, you can click on
            Forgot Password and an e-mail will be sent to the address on record. Alternatively, you can sign in as a new user.	
		    <%ElseIf sOutPut = "FailedAuthChange" Then %>		
			We're sorry, but your combination of e-mail address and password is not recognized. Please retype them.
		    <%ElseIf sOutPut = "Email" Then %>
			If you're an existing customer (you've made a purchase or saved something to your order previously), then we can retrieve your password if you type in the e-mail address on record.
		    <%ElseIf sOutPut = "ChangeCart" Then %>	
			Please enter the e-mail address and password corresponding to an existing saved
            order. Your current saved order can be accessed through the change
            order option on the saved order page.
		    <%ElseIf sOutPut = "EmailMatch" Then %>	
			A matching e-mail address has been found. Please either create a new account, or click on
            Forgot Password for more help
		    <%Else%>
		    Please choose a login with an e-mail account and a password. This will be used for future retrieval of billing and shipping records.
		    <%End If %>	
		    </td>
		  </tr>
		  <tr>
		    <td class="tdContent" align="center"><br />
		   
			  <%If (sOutput <> "FailedAuth" and sOutPut <> "HasID" and sOutPut <> "ChangeCart" and sOutPut <> "FailedAuthChange") Then
			
				If custID_cookie = "" Or sOutPut = "NewAccount" Or sOutput = "EmailMatch" Or sOutPut <> "SameUser" Then %>		   
				<form method="post" action="login.asp" name="login" onSubmit="<%=sSubmitAction%>">
		          <table border="0" width="75%">
		            <tr>
		              <td width="100%" align="center" class="tdBottomTopBanner2">
		       	       
			            <% If sOutPut = "NewAccount" or sOutput = "EmailMatch" Then %>
					New Account 		
		                <% Else  %>
					First Time Here?      
		                <% End If %>
		        
		              </td>
		            </tr>
		            <tr>
		              <td width="100%" align="center" class="tdContent">
		                <table border="0" width="100%">
		                  <tr>
		                    <td width="50%" align="right"><b>E-Mail Address:</b></td>
		                    <td width="50%">
                            <input type="text" name="Email" title="Email Address" class="formDesign" maxlength="100">
		                    </tr>
		                    <tr>
		                      <td width="50%" align="right"><b>Password:</b></td>
		                      <td width="50%">
                              <input type="Password" name="Passwd" title="Password" class="formDesign" maxlength="10"></td>
		                    </tr>
		                    <tr>
		                      <td width="100%" align="center" colspan="2">
					            <input type="image" class="inputImage" src="<%= C_BTN12 %>" name="SignUp" onsubmit="javascript:CheckLoginInput(this)">   
		                      </td>
		                    </tr>
		                  </table>
		                </td>
		              </tr>
		            </table>
		          </form>
		          <br />
		    
					<%If sOutput = "EmailMatch" Then %>
						<form method="post" action="login.asp" name="login" onSubmit="<%= sSubmitAction%>">
					       <table border="0" width="75%">
					        <tr>
					          <td width="100%" align="center" class="tdBottomTopBanner2">Existing Member Login</td>
					        </tr>
					        <tr>
					          <td width="100%" align="center" class="tdContent">
					            <table border="0" width="100%">
						           <tr>
						            <td width="50%" align="right"><b> E-Mail:</b></td>
						            <td width="50%">
                                    <input name="Email" type="text" title="Email Address" class="formDesign" maxlength="100" SIZE="20"></td>
						            </tr>
						            <tr>
						              <td width="50%" align="right"><b>Password:</b></td>
						              <td width="50%">
                                      <input type="password" name="Passwd" title="Password" class="formDesign" maxlength="10" SIZE="20"></td>
						            </tr>
						            <tr>
						              <td width="100%" align="center" colspan="2">
									    <input type="image" class="inputImage" src="<%= C_BTN16 %>" name="Return">
						              </td>
						            </tr>
						          </table>
						      </td>
						      </tr>
						      </table>
						      <p>
						      <a href="login.asp?FPWD=True"><img src="<%= C_BTN17 %>" border="0"></a>
						      </form>
						  <%End If
						
						' End cookies check if
						End If%>
		    
		                  <%Else  %>
		                  <form method="post" action="login.asp" name="login" onSubmit="<%= sSubmitAction%>">
		                    <table border="0"  width="75%">
		                      <tr>
		                        <td width="100%" align="center" class="tdBottomTopBanner2">
		                          <%If sOutPut = "ChangeCart" Then %>	
					Existing Cart Log In
				                  <%Else%>		
					Existing Member
		                          <%End If%>
		                        </td>
		                      </tr>
		                      <tr>
		                        <td width="100%" align="center" class="tdContent">
		                          <table border="0" width="100%">
		                            <tr>
		                              <td width="50%" align="right"><b> E-Mail:</b></td>
		                              <td width="50%">
                                      <input name="Email" type="text" title="Email Address" class="formDesign" maxlength="100"></td>
		                            </tr>
		                            <tr>
		                              <td width="50%" align="right"><b>Password:</b></td>
		                              <td width="50%">
                                      <input type="password" name="Passwd" title="Password" class="formDesign" maxlength="10"></td>
		                            </tr>
		                            <tr>
		                              <td width="100%" align="center" colspan="2">
		                                <%If sOutPut = "ChangeCart" or sOutPut = "FailedAuthChange" Then %>	
						                <input type="image" class="inputImage" src="<%= C_BTN11 %>" name="Change">
		                                <%Else%>
						                <input type="image" class="inputImage" src="<%= C_BTN16 %>" name="Return">
					                    <%End If%>
		                              </td>
		                            </tr>
		                          </table>
		                        </td>
		                      </tr>
		                    </table>
		                    <p><a href="login.asp?FPWD=True"><img src="<%= C_BTN17 %>" border="0"></a>
				            <%If sOutPut = "FailedAuth" or sOutPut = "HasID" Then %>	
					        <a href="login.asp?New=true"><img src="<%= C_BTN19 %>" border="0"></a>
		                    <%End If%>	
		                    </form>
		                    <%End If%>
					</td>
		          </tr>
        
                  <%  
    Else ' Send Email or print confirmation of sent email
    %>
    
 
	                <tr>
	                  <td class="tdContent" align="center"><br />
		                <%
		If sOutput = "Email" Then
		%>	
		                <form method="post" action="login.asp" onSubmit="<%= sSubmitAction%>">
                          <table border="0" width="75%">
                            <tr>
                              <td width="100%" align="center" class="tdBottomTopBanner2">
			Please Type In Your E-Mail Address	
                              </td>
                            </tr>
                            <tr>
                              <td width="100%" align="center" class="tdContent">
                                <table border="0" width="100%">
                                  <tr>
                                    <td width="50%" align="right"><b>E-Mail Address:</b></td>
                                    <td width="50%">
                                    <input type="text" name="Email" title="Email Address" class="formDesign">
                                    </tr>
                                    <tr>
                                      <td width="100%" align="center" colspan="2">
				                        <input type="image" class="inputImage" src="<%= C_BTN18 %>" name="SendEmail">
                                      </td>         
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                            </table>
                          </form>
                        </td>
                      </tr> 
        
		              <% ElseIf sOutPut = "SentEmail" Then %>
			          <table border="0" width="75%">
                        <tr>
                          <td width="100%" align="center" class="tdContent">
					An e-mail with the customer password has been sent to this e-mail:
					        <br /><%=Request.Form("Email")%>
					        <br />
				          </td>
                        </tr>
                      </table>
                      <br />
            
  		              <% ElseIf sOutPut = "FailedEmail" Then %>
			          <table border="0" width="75%">
                        <tr>
                          <td width="100%" align="center" class="tdContent">
					No record exists for a customer with the following e-mail: 
					        <br /><%=Request.Form("Email")%>
					        <p>Would you like to <a href="login.asp?Type=NewAccount">sign in</a> as a new customer?
				            </td>
                          </tr>
                        </table>
	                    <%	
		' End sOutPut If 
		End If
    ' End Send Email If  
    End If
    '-----------------------------------------------------------
    ' End OutPut Block
    '-----------------------------------------------------------
    %>
                      </table>
                    </td>
                  </tr>
              </table>
<!--webbot bot="PurpleText" preview="End Content Section" -->
<!--#include file="templateBottom.asp"-->
</body>
</html>
<%
Call cleanup_dbconnopen
%>
