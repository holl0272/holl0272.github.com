<html>
<body>
<h1>Express Checkout</h1>
<p>
	<b>Note: You must be logged into Developer Central for the sample. </b>
	<a href="https://developer.paypal.com/" target="devcentral">Developer Central</a>
</p>
<p><b>Tip:</b> Check the Log Me in Automatically checkbox in the log in page of 
	Developer Cetnral so you don't have to log in everytime.</p>
<form action="ReviewOrder.asp?ActionCodeType=<%=Request.QueryString("ActionCodeType")%>" method="post">
<table>
	<tr>
		<td>Express Checkout Amount: </td>
		<td><input type="text" name="paymentAmount" value="10.00"></td>
		<td>
							<select name="CurrencyCodeType" ID="Select1">
								<option value="USD">USD</option>
								<option value="GBP">GBP</option>
								<option value="EUR">EUR</option>
								<option value="JPY">JPY</option>
								<option value="CAD">CAD</option>
								<option value="AUD">AUD</option>
							</select>
		</td>
	</tr>
	<tr>
		<td colspan="3">
<input type=image src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" align="left" style="margin-right:7px;" ID="Image1" NAME="Image1">
		</td>
	</tr>
</table>
			<br><b><a href="../Calls.asp">Home</a><b>
</form>
</body>
</html>
