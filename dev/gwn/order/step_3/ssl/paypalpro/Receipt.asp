<!--#include virtual="/PayPalClassicASPSamples/paypal.asp"-->
<!--#include virtual="/PayPalClassicASPSamples/paypal-util.asp"-->

<html>
<body>
<h1>Direct Payment Receipt</h1>

<%
	Dim paypal
	
	If IsEmpty(paypal) Then Set paypal = New PayPalAPI
	
	With paypal
		.DoDirectPayment Request.Form("paymentAmount"), _
			Request.Form("buyerLastName"), Request.Form("buyerFirstName"), _
			Request.Form("buyerAddress1"), Request.Form("buyerAddress2"), _
			Request.Form("buyerCity"), Request.Form("buyerState"), Request.Form("buyerZipCode"), _
			Request.Form("creditCardType"), Request.Form("creditCardNumber"), Request.Form("CVV2"), _
			Request.Form("expMonth"), Request.Form("expYear"), Session("ActionCodeType")
		If IsSuccessful(.pp_caller.Response.Ack) Then
%>
<b>Thank you for your purchase!</b><br><br>
Transaction Details:<br>
<table>
	<tr>
		<td>Transaction ID: </td> 
		<td><%Response.Write .pp_caller.Response.TransactionID%></td>
	</tr>
	<tr>
		<td>AVS Code: </td>
		<td><%Response.Write .pp_caller.Response.AVSCode%></td>
	</tr>
	<tr>
		<td>CVV2 Code: </td>
		<td><%Response.Write .pp_caller.Response.CVV2Code%></td>
	</tr>
	<tr>
		<td>Amount: </td>
		<td><%
			Response.Write GetCurrency(.pp_caller.Response.Amount.currencyID)
			Response.Write .pp_caller.Response.Amount.Value
			%></td>
	</tr>
</table>
<%		
		Else
			PrintErrorMessages(.pp_caller.Response.Errors)
		End If
	End With
%>
			<br><b><a href="../Calls.asp">Home</a><b>

</body>
</html>
