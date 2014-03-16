<!--#include virtual="/PayPalClassicASPSamples/paypal.asp"-->
<!--#include virtual="/PayPalClassicASPSamples/paypal-util.asp"-->

<html>
<body>
<h1>Review Order</h1>

<%
	Dim paypal
	Dim token 

	token = Request.QueryString("token")
			
	If IsEmpty(paypal) Then Set paypal = New PayPalAPI
	
	'Dim util
	'Set util = CreateObject("com.paypal.sdk.COMNetInterop.COMUtil")
	
	With paypal
		If IsEmpty(token) Then
			Dim ReturnURL, CancelURL
			ReturnURL = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("PATH_INFO") & _
				"?paymentAmount=" & Request.Form("paymentAmount") & "&CurrencyCodeType=" + Request.Form("CurrencyCodeType") & "&ActionCodeType=" + Request.QueryString("ActionCodeType")
			CancelURL = Replace(ReturnURL, "ReviewOrder", "ExpressCheckout")
			.SetExpressCheckout Request.Form("paymentAmount"), ReturnURL, CancelURL, Request.QueryString("ActionCodeType"), Request.Form("CurrencyCodeType")
			
			If IsSuccessful(.pp_caller.Response.Ack) Then
				token = .pp_caller.Response.Token
				Response.Redirect( "https://www."  & ENVIRONMENT & ".paypal.com/cgi-bin/webscr?cmd=_express-checkout&token=" & token)
			Else
				PrintErrorMessages(.pp_caller.Response.Errors)
			End If
		Else
			.GetExpressCheckoutDetails token 
			
			If IsSuccessful(.pp_caller.Response.Ack) Then
			%>
<b>Shipping Address:</b><br>
			<%
				With .pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo
					Response.Write .PayerName.FirstName & " " & .PayerName.LastName & "<br>"
					Response.Write .Address.Street1 & "<br>"
					Response.Write .Address.CityName & ", " & .Address.StateOrProvince & " " & .Address.PostalCode & "<br>"
					Response.Write .Address.CountryName
				End With
			%>
<br><br><a href="ECReceipt.asp?token=<%=.pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.Token%>
	&payerID=<%=.pp_caller.Response.GetExpressCheckoutDetailsResponseDetails.PayerInfo.PayerID%>
	&paymentAmount=<%=Request.QueryString("paymentAmount")%>&ActionCodeType=<%=Request.QueryString("ActionCodeType")%>&CurrencyCodeType=<%=Request.QueryString("CurrencyCodeType")%>">Pay</a>
			<%
			Else
				PrintErrorMessages(.pp_caller.Response.Errors)
			End If	
		End If
	End With
%>
			<br>
			<b><a href="../Calls.asp">Home</a></b>

</body>
</html>
