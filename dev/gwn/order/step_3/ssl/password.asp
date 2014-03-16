<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->

<!--#include file ="SFLib/incProcOrder.asp"-->
<!--#include file="SFLib/incLogin.asp"-->
	<%
	'Const vDebug = 0
    
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4014.0.2

    '@FILENAME: password.asp
	 



	'@DESCRIPTION: Handles password Information

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO

Dim sStatus, sOutput, sEmail, sPasswd, sOldPassword, sNewPassword, iAuthenticate, bEmailAddress
sStatus = Trim(Request.QueryString("status"))
	
If sStatus = "fpwd" then
	sOutput = "FPWD"
End If	
If Trim(Request.Form("SendEmail.x")) <> "" or (Request.Form.Count=1 and Trim(Request.Form("SubmitNewPWD.x"))="") Then
'above if is there to see if this is someone requesting a new password,
'if they hit return while in the e-mail box, the first condition is false,
'so I had to add the other conditions
	sEmail = Trim(Request.Form("Email"))
	sPasswd = Trim(Request.Form("Password"))

	bEmailAddress = SendPassword(sEmail)
			
	If bEmailAddress = 1 Then
		sOutPut = "EmailSent"
	Else
		sOutPut = "NoEmailMatch"
	End If	

End If


%>
<html>
<head>
<title><%= C_STORENAME %>- Password Retrieval Page</title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" href="include_commonElements/styles.css" type="text/css">

<script language="javascript" src="SFLib/common.js" type="text/javascript"></script>
<script language="javascript" src="SFLib/sfCheckErrors.js"></SCRIPT>
</head>

<body <%= mstrBodyStyle %>>

<!--#include file="templateTop.asp"-->
<!--webbot bot="PurpleText" preview="Begin Content Section" -->
<table border="0" cellspacing="0" cellpadding="0" id="tblMainContent">
        <tr>
          <td align="center" class="tdMiddleTopBanner">
	<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="100%" align="center">
        <tr>
          <td align="center" class="tdMiddleTopBanner"><h3>
   		    <% If sOutput = "ChangePwd" Then %>
		Change Login/Password
		    <% ElseIf sOutput = "EmailSent" Then %>
		E-Mail With Password Sent
		    <% ElseIf sOutput = "NoEmailMatch" Then %>
		No Matching E-Mail Found	
		    <% ElseIf sOutput = "FPWD" Then %>
		Forgotten Password Help		
		    <% End If %>
	        </h3></td>
	        </tr>	
            <tr>
	          <td class="tdBottomTopBanner2" align="center">
		        <% If sOutput = "EmailSent" Then %>
		An e-mail was sent to the address.
		        <% ElseIf sOutput = "NoEmailMatch" Then %>
                No e-mail address was found to match.
		        <% ElseIf sOutput = "FPWD" Then %>
		Please enter your e-mail address and your password will be sent to you immediately via e-mail.
		<% End If %>
		</td>
	            </tr>	
	
	                        <% If sOutput="FPWD" Then %>
		                    <tr>
		                      <td class="tdContent" align="center"><br />
			                    <form method="post" name=thisForm>
				                  <table border="0" width="100%">
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
							                <td width="50%"><input type="text" name="Email" title="Email Address" class="formDesign">
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
        
	                          <% ElseIf sOutput = "EmailSent" Then %>
		                      <tr align="center">			      			 
		                        <td>
			                      <br />
			                      <table border="0" width="100%" align="center">
                                    <tr align="center">
                                      <td width="100%" align="center" class="tdContent">
					An e-mail with the customer password has been sent to this e-mail:
					                    <br /><%= sEmail %>
					                    <br />
				                      </td>
                                    </tr>
                                  </table>
                                  <br />
                                </td>
                              </tr>    
                              <% ElseIf sOutPut = "NoEmailMatch" Then %>
                              <tr>
			                    <table border="0" width="100%">
                                  <tr>
                                    <td width="100%" align="center" class="tdContent">
					No record exists for a customer with the following e-mail: 
					                  <br /><%= sEmail %>
					                  <br />Please try again.
				                    </td>
                                  </tr>
                                </table>
                                <br />
            
                              </tr>   
                              <% 
	   End If 
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
