<!--#include virtual="/PayPalClassicASPSamples/paypal.asp"-->
<!--#include virtual="/PayPalClassicASPSamples/paypal-util.asp"-->

<html>
<body>
<h1>Express Checkout Receipt</h1>

<%
	Dim paypal
	
	If IsEmpty(paypal) Then Set paypal = New PayPalAPI
	
	With paypal
		.DoExpressCheckoutPayment Request.QueryString("token"), Request.QueryString("payerID"), Request.QueryString("paymentAmount"), Request.QueryString("ActionCodeType"), Request.QueryString("CurrencyCodeType")
		If IsSuccessful(.pp_caller.Response.Ack) Then
%>
<b>Thank you for your purchase!</b><br><br>
Transaction Details:<br>
<table>
	<tr>
		<td>Transaction ID: </td> 
		<td><%Response.Write .pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.TransactionID%></td>
	</tr>
	<tr>
		<td>Amount: </td>
		<td><%
			Response.Write GetCurrency(.pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.GrossAmount.currencyID)
			Response.Write .pp_caller.Response.DoExpressCheckoutPaymentResponseDetails.PaymentInfo.GrossAmount.Value
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
