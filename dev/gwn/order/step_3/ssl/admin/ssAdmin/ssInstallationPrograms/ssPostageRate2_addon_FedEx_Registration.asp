<%Option Explicit 
'********************************************************************************
'*   Postage Rate Administration						                        *
'*   Release Version: 2.0			                                            *
'*   Release Date: September 15, 2002											*
'*                                                                              *
'*   Release Notes:                                                             *
'*                                                                              *
'*   The contents of this file are protected by United States copyright laws    *
'*   as an unpublished work. No part of this file may be used or disclosed      *
'*   without the express written permission of Sandshot Software.               *
'*                                                                              *
'*   (c) Copyright 2002 Sandshot Software.  All rights reserved.                *
'********************************************************************************

sub debugprint(sField1,sField2)
Response.Write "<H3>" & sField1 & ": " & sField2 & "</H3><BR>"
end sub

dim conn,strConn,strDBpath,sql
dim mrsTest
dim mblnUpgraded
dim mblnValidConnection
dim mstrAction
dim mblnError
Dim mstrMessage
Dim pstrFedExURL
Dim pstrData
Dim pstrResponseText
Dim pstrDSN

	mblnError = False
	mstrAction = Request.Form("Action")
	
	pstrDSN = Application("DSN_NAME")
	If Len(pstrDSN) = 0 Then pstrDSN = Session("DSN_NAME")

	if len(mstrAction) > 0 then

		Set conn = Server.CreateObject ("ADODB.Connection")
		conn.Open pstrDSN
	End If
	If mstrAction = "register" Then

		pstrFedExURL = Request.Form("FedExURL")
		pstrFedExURL = "gatewaybeta.fedex.com:443"
		pstrFedExURL = "gateway.fedex.com:443"
		
		pstrData = "0," & (Chr(34) & "211" & Chr(34)) _
				  & "1," & (Chr(34) & "Subscribe" & Chr(34)) _
				  & "4003," & (Chr(34) & Trim(Request.Form("a4003")) & Chr(34)) _
				  & "4007," & (Chr(34) & Trim(Request.Form("a4007")) & Chr(34)) _
				  & "4011," & (Chr(34) & Trim(Request.Form("a4011")) & Chr(34)) _
				  & "4008," & (Chr(34) & Trim(Request.Form("a4008")) & Chr(34)) _
				  & "4015," & (Chr(34) & Trim(Request.Form("a4015")) & Chr(34)) _
				  & "4012," & (Chr(34) & Trim(Request.Form("a4012")) & Chr(34)) _
				  & "10," & (Chr(34) & Trim(Request.Form("a10")) & Chr(34)) _
				  & "4013," & (Chr(34) & Trim(Request.Form("a4013")) & Chr(34)) _
				  & "4014," & (Chr(34) & Trim(Request.Form("a4014")) & Chr(34)) _
				  & "99," & (Chr(34) & "" & Chr(34))

		'debugprint "pstrData",pstrData

		'On Error Resume Next

			If Err.number <> 0 Then	Err.Clear
			Dim pobjXMLHTTP
			Set pobjXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
			With pobjXMLHTTP
				.open "POST", "https://" & pstrFedExURL & "/GatewayDC", False
				.setRequestHeader "Referer", Trim(Request.Form("a4007"))
				.setRequestHeader "Host", pstrFedExURL
				.setRequestHeader "Accept","image/gif, image/jpeg, image/pjpeg, text/plain,text/html, */*"
				.setRequestHeader "Content-Type","image/gif"
				.setRequestHeader "Content-Length", CStr(Len(pstrData))
				.send pstrData
				pstrResponseText  = .responseText
				'Response.Write "() Error:" & Err.number & " - " & Err.Description & " (" & Err.Source & ")" & "<BR>"	
				pstrResponseText = Replace(pstrResponseText,"99,""""","")
				'Response.Write "pstrResponseText =" & pstrResponseText & "<br>" & vbcrlf
				
				'Now split out the results
				Dim pblnEven
				Dim plngPos1
				
				plngPos1 = 1
				Do While plngPos1 > 0
					plngPos1 = Instr(plngPos1,pstrResponseText,Chr(34))
					If plngPos1 > 0 Then
						If pblnEven Then
							plngPos1 = plngPos1 + 1
						Else
							pstrResponseText = Left(pstrResponseText,plngPos1 - 1) & Right(pstrResponseText,Len(pstrResponseText)-plngPos1)
						End If
						pblnEven = Not pblnEven
					End If
					'Response.Write "pstrResponseText =" & pstrResponseText & "<br>" & vbcrlf
				Loop
				
				Dim paryResults
				Dim pstrMeterNumber
				Dim pstrSubscription
				Dim i
				paryResults = Split(pstrResponseText,Chr(34))
				
				For i = 0 To UBound(paryResults)-1
					paryResults(i) = Split(paryResults(i),",")
				Next 'i

				For i = 0 To UBound(paryResults)-1
					'Response.Write paryResults(i)(0) & "=" & paryResults(i)(1) & "<br>" & vbcrlf
					'Response.Flush
					Select Case paryResults(i)(0)
						Case "498"
							pstrMeterNumber = paryResults(i)(1)
						Case "3"
							mstrMessage = "<H4><font color='Red'>" & paryResults(i)(1) & "</FONT></H4>" & vbcrlf
						Case Else
							If Instr(paryResults(i)(0),"4021") > 0 Then
								pstrSubscription = pstrSubscription & paryResults(i)(1) & "<BR>"
							End If
					End Select
				Next 'i

			End With
			set pobjXMLHTTP = nothing

	End If

	'Test Connection to the database
	If len(pstrDSN) > 0 then
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open pstrDSN
		If (conn.State = 1) Then
			mblnValidConnection = True
		Else
			mblnValidConnection = False
			mstrMessage = "<H3><Font Color='Red'>Could not connect to the database. Error: " & Err.number & " - " & Err.Description & "</FONT></H3>"
			Err.Clear
		End If
	Else
		mblnValidConnection = False
		mstrMessage = "<H3><Font Color='Red'>Could not connect to the database</FONT></H3>"
	End If

	'Test to see if DB has already been upgraded
	If mblnValidConnection Then
		sql = "Select ssShippingCarrierID from ssShippingCarriers"
		Set mrsTest = conn.Execute(sql)

		mblnUpgraded = (Err.number=0)
		mrsTest.Close
		Set mrsTest = Nothing
		Err.Clear
		
	Else
		mblnUpgraded = False
	End If

	On Error Resume Next
	
	If Len(pstrMeterNumber) > 0 Then
		conn.Execute "Update ssShippingCarriers Set ssShippingCarrierPassword='" & pstrMeterNumber & "' Where ssShippingCarrierID=4",,128
	End If
	
	conn.Close 
	Set conn = Nothing
	
%><html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>FedEx API Registration</title>
</head>

<body>
<h2>Postage Rate Component Installation Guide for StoreFront 5.0</h2>

<h3>FedEx Registration</h3>

<p>In order to obtain FedEx real time rate quotes you must have an account with 
FedEx and register to use their services.</p>

<p>The Process:</p>

<ol>
  <li>Obtain a FedEx account number. If you do not already have one click <a href="http://www.fedex.com/us/customer/openaccount/?link=2" target="_blank">here</a>.</li>
  <li>Once you have your account number you must notify FedEx of your desire to use their real-time rates. Register <a href="https://www.fedex.com/globaldeveloper/shipapi/register.html?link=2" target="_blank">here</a> to do so.
  <ul>
  <li><em>Type of Business</em>: - Corporate Developer</li>
  <li><em>Communication path to FedEx</em>: - FSM Direct</li>
  <li><em>Data Format</em>: - XML Tools</li>
  </ul>
  </li>
  <li>Once you receive notification from FedEx saying you can use there real-time rates you must register your account on their server to obtain a meter number - that is what this page does</li>
</ol>

<p>&nbsp;</p>



<% If mblnUpgraded Then 

	If Len(pstrMeterNumber) > 0 Then
%>
	<h3>You have been successfully registered</h3>
	
	<h4>Your meter number is: <%= pstrMeterNumber %></h4>
	<h4>You are eligible for the following services: <%= pstrSubscription %></h4>
	
	<p>Your meter number has been automatically saved to the database.</p>
<%
	Else
		Response.Write mstrMessage
%>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.a4003.value == "")
  {
    alert("Please enter a value for the \"Your Name\" field.");
    theForm.a4003.focus();
    return (false);
  }

  if (theForm.a4007.value == "")
  {
    alert("Please enter a value for the \"Company Name\" field.");
    theForm.a4007.focus();
    return (false);
  }

  if (theForm.a4008.value == "")
  {
    alert("Please enter a value for the \"street address\" field.");
    theForm.a4008.focus();
    return (false);
  }

  if (theForm.a4011.value == "")
  {
    alert("Please enter a value for the \"City\" field.");
    theForm.a4011.focus();
    return (false);
  }

  if (theForm.a4012.value == "")
  {
    alert("Please enter a value for the \"State\" field.");
    theForm.a4012.focus();
    return (false);
  }

  if (theForm.a4013.value == "")
  {
    alert("Please enter a value for the \"postal code\" field.");
    theForm.a4013.focus();
    return (false);
  }

  if (theForm.a4014.value == "")
  {
    alert("Please enter a value for the \"Country Code\" field.");
    theForm.a4014.focus();
    return (false);
  }

  if (theForm.a4014.value.length < 2)
  {
    alert("Please enter at least 2 characters in the \"Country Code\" field.");
    theForm.a4014.focus();
    return (false);
  }

  if (theForm.a4014.value.length > 2)
  {
    alert("Please enter at most 2 characters in the \"Country Code\" field.");
    theForm.a4014.focus();
    return (false);
  }

  if (theForm.a4015.value == "")
  {
    alert("Please enter a value for the \"phone number\" field.");
    theForm.a4015.focus();
    return (false);
  }

  if (theForm.a4015.value.length < 10)
  {
    alert("Please enter at least 10 characters in the \"phone number\" field.");
    theForm.a4015.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.a4015.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"phone number\" field.");
    theForm.a4015.focus();
    return (false);
  }

  if (theForm.a10.value == "")
  {
    alert("Please enter a value for the \"FedEx Account Number\" field.");
    theForm.a10.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.a10.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"FedEx Account Number\" field.");
    theForm.a10.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="ssPostageRate2_addon_FedEx_Registration.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" id="FrontPage_Form1" language="JavaScript">
	<input TYPE="hidden" NAME="action" VALUE="register">
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td width="50%" height="19">Your name.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="Your Name" b-value-required="TRUE" --><input type="text" name="a4003" size="50" value="<%= Request.Form("a4003") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your company’s name.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="Company Name" b-value-required="TRUE" --><input type="text" name="a4007" size="50" value="<%= Request.Form("a4007") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your street address.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="street address" b-value-required="TRUE" --><input type="text" name="a4008" size="50" value="<%= Request.Form("a4008") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your city name.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="City" b-value-required="TRUE" --><input type="text" name="a4011" size="50" value="<%= Request.Form("a4011") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="18">Your state abbreviation.</td>
      <td width="50%" height="18">
      <!--webbot bot="Validation" s-display-name="State" b-value-required="TRUE" --><input type="text" name="a4012" size="5" value="<%= Request.Form("a4012") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your zip code.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="postal code" b-value-required="TRUE" --><input type="text" name="a4013" size="8" value="<%= Request.Form("a4013") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your Country Code</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="Country Code" b-value-required="TRUE" i-minimum-length="2" i-maximum-length="2" --><input type="text" name="a4014" size="4" value="US" maxlength="2" value="<%= Request.Form("a4014") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your phone number, no dashes.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="phone number" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="10" --><input type="text" name="a4015" size="15" value="<%= Request.Form("a4015") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">Your FedEx account number.</td>
      <td width="50%" height="19">
      <!--webbot bot="Validation" s-display-name="FedEx Account Number" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" --><input type="text" name="a10" size="15" value="<%= Request.Form("a10") %>"></td>
    </tr>
    <tr>
      <td width="50%" height="19">&nbsp;</td>
      <td width="50%" height="19"><input type="submit" value="Submit" name="B1"></td>
    </tr>
  </table>
</form>
<%	End If	'mblnUpgraded %>
<% Else %>
<h2><font color="red">Error: Your database has not been upgraded to use the Postage Rate Add-on</font></h2>
<h3><a href="../ssHelpFiles/ssInstallationPrograms/ssMasterDBUpgradeTool.asp?UpgradeItem=PostageRate">Click here to upgrade your database</a></h3>
<% End If	'mblnUpgraded %>


</body>

</html>