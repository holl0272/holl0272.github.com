<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include File = "processor.asp"-->
<%

On Error Goto 0

'--------------------------------------------------------------------------
' Orbital Function
'--------------------------------------------------------------------------
Dim amount
Dim cardName
Dim cardCompany
Dim cardAddr1
Dim cardAddr2
Dim cardCity
Dim cardState
Dim cardZIP
Dim cardCountry
Dim cardNumber
Dim cardExpMonth
Dim cardExpYear
Dim cardCCV
Dim pstrResult
Dim ProcessorToTest

If Len(Request.Form) > 0 Then
	amount = Trim(Request.Form("amount"))
	cardName = Trim(Request.Form("cardName"))
	cardCompany = Trim(Request.Form("cardCompany"))
	cardAddr1 = Trim(Request.Form("cardAddr1"))
	cardAddr2 = Trim(Request.Form("cardAddr2"))
	cardCity = Trim(Request.Form("cardCity"))
	cardState = Trim(Request.Form("cardState"))
	cardZIP = Trim(Request.Form("cardZIP"))
	cardCountry = Trim(Request.Form("cardCountry"))
	cardNumber = Trim(Request.Form("cardNumber"))
	cardExpMonth = Trim(Request.Form("cardExpMonth"))
	cardExpYear = Trim(Request.Form("cardExpYear"))
	cardCCV = Trim(Request.Form("cardCCV"))
	ProcessorToTest = Trim(Request.Form("ProcessorToTest"))
	iOrderID = Trim(Request.Form("iOrderID"))

	'Declare some global variables
	Dim sPaymentServer:		sPaymentServer = ""
	Dim sGrandTotal:		sGrandTotal = amount
	Dim sCustCardNumber:	sCustCardNumber = cardNumber
	Dim sCustCardExpiry:	sCustCardExpiry = cardExpMonth & "/" & cardExpYear
	Dim sCustCardExpiryMonth:	sCustCardExpiryMonth = cardExpMonth
	Dim sCustCardExpiryYear:	sCustCardExpiryYear = cardExpYear
	Dim mstrPayCardCCV:		mstrPayCardCCV = cardCCV
	Dim iOrderID:
	Dim sCustCardName:		sCustCardName = cardName
	Dim sCustCompany:		sCustCompany = cardCompany
	Dim sCustAddress1:		sCustAddress1 = cardAddr1
	Dim sCustAddress2:		sCustAddress2 = cardAddr2
	Dim sCustCity:			sCustCity = cardCity
	Dim sCustState:			sCustState = cardState
	Dim sCustZip:			sCustZip = cardZIP
	Dim sCustCountry:		sCustCountry = cardCountry
	Dim sCustPhone:			sCustPhone = ""
	Dim sCustEmail:			sCustEmail = ""
	Dim sCustFax:			sCustFax = ""

	Dim aryCardName:		aryCardName = Split(cardName, " ")
	Dim sCustFirstName:		sCustFirstName = aryCardName(0)
	Dim sCustLastName:		sCustLastName = aryCardName(1)

	Dim sShipCustFirstName:		sShipCustFirstName = sCustFirstName
	Dim sShipCustLastName:		sShipCustLastName = sCustLastName
	Dim sShipCustCity:			sShipCustCity = sCustCity
	Dim sShipCustState:			sShipCustState = sCustState
	Dim sShipCustZip:			sShipCustZip = sCustZip
	Dim sShipCustCountry:		sShipCustCountry = sCustCountry

	Dim sLogin:				sLogin = ""
	Dim sPassword:			sPassword = ""
	Dim sMercType:			sMercType = "AUTHCAPTURE"	'AUTHONLY
	Dim iProcResponse
	Dim ProcResponse
	Dim ProcErrMessage
	Dim ProcMessage
	Dim ProcActionCode
	Dim ProcResponseCode
	Dim ProcAuth
	Dim ProcAuthCode
	Dim ProcRefCode
	Dim ProcCustNumber
	Dim ProcAvsCode
	Dim ProcAvsMsg
	Dim sTransMethod
	
	'Adapter code for SignioPayProFlow
	Dim aReferer(3): aReferer(2) = Request.ServerVariables("REMOTE_ADDR")
		
	'Adapter code for electricblanket.net
	Dim sCustCardCVV: sCustCardCVV = cardCCV

	'Now run test
	Select Case ProcessorToTest
		Case "Orbital":
			sTransMethod = "19"	'varies as this is a non-standard processor
			pstrResult = Orbital(1)
		Case "LinkPoint":
			sTransMethod = "7"
			pstrResult = LinkPoint(0)
		Case "PayFlowPro30":
			sTransMethod = "17"
			pstrResult = SignioPayProFlow(0)
		Case "Verisign PayFlow Pro":
			sTransMethod = "3"
			pstrResult = SignioPayProFlow(0)
	End Select	
	
Else
	iOrderID = 0
	amount = "9.99"
	cardName = "Rick Dennery"
	cardAddr1 = "1958 Brooke Farm Ct"
	cardAddr2 = ""
	cardCity = "Woodbridge"
	cardState = "VA"
	cardZIP = "22192"
	cardCountry = "US"
	cardNumber = "4111111111111111"
	cardExpMonth = "05"
	cardExpYear = "06"
	cardCCV = ""
End If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Processor Test</title>
</head>

<body>

<form method="POST" action="processor_Test.asp">

<TABLE cellSpacing="0" cellPadding="2" border="1" style="border-collapse: collapse" bordercolor="#000000">
	<TR>
		<TD>Order Number</TD>
		<TD><input type="text" name="iOrderID" size="20" value="<%= iOrderID + 1 %>" ID="iOrderID"></TD>
	</TR>
	<TR>
		<TD>Amount</TD>
		<TD><input type="text" name="amount" size="20" value="<%= amount %>" ID="amount"></TD>
	</TR>
	<tr>
		<TD>Card Name</TD>
		<TD><input type="text" name="cardName" size="20" value="<%= cardName %>" ID="cardName"></TD>
	</tr>
	<tr>
		<TD>Company</TD>
		<TD><input type="text" name="cardCompany" size="20" value="<%= cardCompany %>" ID="cardCompany"></TD>
	</tr>
	<tr>
		<TD>Addr 1</TD>
		<TD>
		<input type="text" name="cardAddr1" size="20" value="<%= cardAddr1 %>"></TD>
	</tr>
	<tr>
		<TD>Addr 2</TD>
		<TD><input type="text" name="cardAddr2" size="20" value="<%= cardAddr2 %>"></TD>
	</tr>
	<tr>
		<TD>City</TD>
		<TD><input type="text" name="cardCity" size="20" value="<%= cardCity %>"></TD>
	</tr>
	<tr>
		<TD>State</TD>
		<TD><input type="text" name="cardState" size="20" value="<%= cardState %>"></TD>
	</tr>
	<tr>
		<TD>ZIP</TD>
		<TD><input type="text" name="cardZIP" size="20" value="<%= cardZIP %>"></TD>
	</tr>
	<TR>
		<TD>Country</TD>
		<TD><input type="text" name="cardCountry" size="20" value="<%= cardCountry %>"></TD>
	</TR>
	<TR>
		<TD>Card Number</TD>
		<TD>
		<input type="text" name="cardNumber" size="20" value="<%= cardNumber %>"></TD>
	</TR>
	<tr>
		<TD>Card Exp Mo</TD>
		<TD><input type="text" name="cardExpMonth" size="20" value="<%= cardExpMonth %>" ID="Text1"></TD>
	</tr>
	<tr>
		<TD>Card Exp Yr</TD>
		<TD><input type="text" name="cardExpYear" size="20" value="<%= cardExpYear %>" ID="Text2"></TD>
	</tr>
	<tr>
		<TD>CCV</TD>
		<TD><input type="text" name="cardCCV" value="<%= cardCCV %>"></td>
    </tr>
    <tr>
      <td>Processor</td>
      <td>
        <input type="radio" name="ProcessorToTest" id="ProcessorToTest0" value="Orbital" <% If ProcessorToTest = "Orbital" Then Response.Write " checked" %>>Orbital<br />
        <input type="radio" name="ProcessorToTest" id="ProcessorToTest1" value="LinkPoint" <% If ProcessorToTest = "LinkPoint" Then Response.Write " checked" %>>LinkPoint<br />
        <input type="radio" name="ProcessorToTest" id="ProcessorToTest2" value="PayFlowPro30" <% If ProcessorToTest = "PayFlowPro30" Then Response.Write " checked" %>>PayFlowPro30<br />
        <input type="radio" name="ProcessorToTest" id="ProcessorToTest3" value="Verisign PayFlow Pro" <% If ProcessorToTest = "Verisign PayFlow Pro" Then Response.Write " checked" %>>Verisign PayFlow Pro<br />
      </td>
    </tr>
	<tr><td></td><td><input type="submit" value="Submit"></td></tr>
	<tr><td colspan="2"><hr></td></tr>
	<tr><td colspan="2"><textarea rows=20 cols=180><%= pstrResult %></textarea></td></tr>
</TABLE>
</form>
<table border="1" cellspacing="1" cellpadding="2" style="border-collapse: collapse" bordercolor="#000000">
	<tr>
		<td>Visa </td>
		<td>4012888888881</td>
	</tr>
	<tr>
		<td>Visa Purchasing Card II</td>
		<td>4055011111111111</td>
	</tr>
	<tr>
		<td>MasterCard </td>
		<td>5454545454545454</td>
	</tr>
	<tr>
		<td>MasterCard Purchasing Card II</td>
		<td>5405222222222226</td>
	</tr>
	<tr>
		<td>American Express </td>
		<td>371449635398431</td>
	</tr>
	<tr>
		<td>Discover </td>
		<td>6011000995500000</td>
	</tr>
	<tr>
		<td>Diners </td>
		<td>36438999960016</td>
	</tr>
	<tr>
		<td>JCB </td>
		<td>3566002020140006</td>
	</tr>
</table>
<p><br />
&nbsp;<br />
<br />
&nbsp;<br />
<br />
<br />
<br />
&nbsp;</p>
</body>

</html>