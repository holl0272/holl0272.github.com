<%
'*******************************************************************************
' Copyright (C) 2006 Google Inc.
' 
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'  
'      http://www.apache.org/licenses/LICENSE-2.0
'  
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'*******************************************************************************

'*******************************************************************************
' Please refer to the Google Checkout ASP Sample Code Documentation
' for requirements and guidelines on how to use the sample code.
'
' This sample code demonstrates how to utilize the Google Checkout ASP Client 
' Library to generate <checkout-shopping-cart> XML (referred below as 
' "Cart XML") and to place a new order with the Cart XML through an 
' HTML POST form.
'*******************************************************************************

' Include ASP libraries used in shopping cart demo
%>

<!--#INCLUDE file="GlobalAPIFunctions.asp"--> 
<!--#INCLUDE file="CheckoutAPIFunctions.asp"--> 

<%

'**********************Create checkout shopping cart **************************'

' Define objects used to create the shopping cart
Dim domTaxArea
Dim domShippingRestrictions

Dim elemItemName
Dim elemItemDescription
Dim elemQuantity
Dim elemUnitPrice
Dim elemTaxTableSelector
Dim elemShippingTaxed
Dim elemRate
Dim elemTaxAreaState
Dim elemTaxAreaCountry
Dim elemTaxAreaZip
Dim elemPrice
Dim elemMerchantCalculationsUrl
Dim elemAcceptMerchantCoupons
Dim elemAcceptGiftCertificates
Dim elemEditCartUrl
Dim elemContinueShoppingUrl
Dim elemMerchantPrivateItemData
Dim elemMerchantPrivateData

Dim attrName
Dim attrMerchantCalculated
Dim attrStandalone
Dim attrAllowedCountryArea
Dim attrExcludedCountryArea

Dim dtmCartExpiration

Dim strAllowedState
Dim strAllowedZip
Dim strExcludedState
Dim strExcludedZip

Dim arrayAllowedState
Dim arrayAllowedZip
Dim arrayExcludedState
Dim arrayExcludedZip

Dim checkoutPostData
Dim diagnoseResponse

' Build XML for items in the shopping cart. The shopping cart
' has the following structure:
' <shopping-cart>
'     <items>
'         <item>
'             <item-name>Dry Food Pack AA1453</item-name>
'             <item-description>This food is very nutritious.</item-description>
'             <quantity>1</quantity>
'             <unit-price currency="USD">35.00</unit-price>
'             <tax-table-selector>food</tax-table-selector>
'             <merchant-private-item-data>
'               <item-note>Product Number N15037124531</item-note>
'             </merchant-private-item-data>
'         </item>
'         <!-- More items may be included using the same XML structure -->
'     </items>
' </shopping-cart>

' The XML for an individual item is created by defining data fields
' for the item and then calling the createItem() function, which
' is in the GlobalAPIFunctions.asp file.

'  * +++ CHANGE ME +++
' You will need to modify calls to functions like createItem,
' addAllowedAreas, addExcludedAreas, createFlatRateShipping and
' numerous others in this file to reflect the items in the
' customer's shopping cart, the shipping options available for
' those items and the tax tables that you use to calculate taxes.

' Specify item data and create an item to include in the order
elemItemName = "Dry Food Pack AA1453" 
elemItemDescription = "A pack of highly nutritious dried food for emergency " _
    & "- store in your garage for up to one year!!"
elemQuantity = "1"
elemUnitPrice = "35.00"
elemTaxTableSelector = "food"
createItem elemItemName, elemItemDescription, elemQuantity, elemUnitPrice, _
    elemTaxTableSelector, elemMerchantPrivateItemData

' Specify item data and create a second item to include in the order
elemItemName = "MegaSound 2GB MP3 Player"
elemItemDescription = "Portable MP3 player - stores 500 songs"
elemQuantity = "1"
elemUnitPrice = "178.00"
elemTaxTableSelector = ""
elemMerchantPrivateItemData = _
    "<item-note>Product Number N15037124531</item-note>"
createItem elemItemName, elemItemDescription, elemQuantity, elemUnitPrice, _
    elemTaxTableSelector, elemMerchantPrivateItemData

' Specify an expiration date for the order and build <shopping-cart>
dtmCartExpiration = "2006-12-31T23:59:59"
elemMerchantPrivateData = _
    "<merchant-note>My order number 9876543</merchant-note>"
createShoppingCart dtmCartExpiration, elemMerchantPrivateData

' Create list of areas where a particular shipping option is available
attrAllowedCountryArea = "ALL"  ' OR: "CONTINENTAL_48", "FULL_50_STATES"
arrayAllowedState = Array()     ' Ex: Array("CA", "NY", "DC", "NC")
arrayAllowedZip = Array()       ' Ex: Array("94043", "94086", "91801", "91362")
Set domShippingRestrictions = addAllowedAreas(attrAllowedCountryArea, _
    arrayAllowedState, arrayAllowedZip)

' Create list of areas where a particular shipping option is not available
attrExcludedCountryArea = ""
arrayExcludedState = Array("AL", "MA", "MT", "WA")
arrayExcludedZip = Array()
Set domShippingRestrictions = addExcludedAreas(attrExcludedCountryArea, _
    arrayExcludedState, arrayExcludedZip)

' Create a <flat-rate-shipping> option with shipping restrictions
attrName = "UPS Ground"
elemPrice = "8.50"
createFlatRateShipping attrName, elemPrice, domShippingRestrictions

' Create a <merchant-calculated-shipping> option without shipping restrictions
attrName = "SuperShip"
elemPrice = "10.00"
domShippingRestrictions = ""
createMerchantCalculatedShipping attrName, elemPrice, domShippingRestrictions

' Create a <pickup> shipping option
attrName = "Pickup"
elemPrice = "0.00"
createPickup attrName, elemPrice

' Create tax tables for the order. Tax tables have the
' following XML structure:
' <tax-tables>
'     <default-tax-table>
'         <tax-rules>
'             <default-tax-rule>
'                 <shipping-taxed>true</shipping-taxed>
'                 <rate>0.0825</rate>
'                 <tax-area>
'                     <!-- could also contain country or zip areas>
'                     <us-state-area>
'                         <state>NY</state>
'                     </us-state-area>
'                 </tax-area>
'             </default-tax-rule>
'         </tax-rules>
'     </default-tax-table>
'     <alternate-tax-tables>
'         <alternate-tax-table>
'             <alternate-tax-rules>
'                 <alternate-tax-rule>
'                     <rate>0.0825</rate>
'                     <tax-area>
'                         <!-- could also contain country or zip areas>
'                         <us-state-area>
'                             <state>NY</state>
'                         </us-state-area>
'                     </tax-area>
'                 </alternate-tax-rule>
'             </alternate-tax-rules>
'         </alternate-tax-table>
'     </alternate-tax-tables>
' </tax-tables>
'
' +++ CHANGE ME +++
' You will need to update the tax tables to match those
' used to calculate taxes for your store

' Build <default-tax-table>
elemRate = "0.0825"
elemTaxAreaCountry = "ALL"
Set domTaxArea = createTaxArea("country", elemTaxAreaCountry)
elemShippingTaxed = "false"
createDefaultTaxRule elemRate, domTaxArea, elemShippingTaxed

elemRate = "0.0800"
elemTaxAreaState = "NY"
Set domTaxArea = createTaxArea("state", elemTaxAreaState)
elemShippingTaxed = "true"
createDefaultTaxRule elemRate, domTaxArea, elemShippingTaxed

' Build an <alternate-tax-table>
elemRate = "0.0225"
elemTaxAreaState = "CA"
Set domTaxArea = createTaxArea("state", elemTaxAreaState)
createAlternateTaxRule elemRate, domTaxArea

elemRate = "0.0200"
elemTaxAreaState = "NY"
Set domTaxArea = createTaxArea("state", elemTaxAreaState)
createAlternateTaxRule elemRate, domTaxArea

attrStandalone = "false"
attrName = "food"
createAlternateTaxTable attrStandalone, attrName


' Build another <alternate-tax-table>
elemRate = "0.0500"
elemTaxAreaCountry = "FULL_50_STATES"
Set domTaxArea = createTaxArea("country", elemTaxAreaCountry)
createAlternateTaxRule elemRate, domTaxArea

elemRate = "0.0600"
elemTaxAreaZip = "9404*"
Set domTaxArea = createTaxArea("zip", elemTaxAreaZip)
createAlternateTaxRule elemRate, domTaxArea

attrStandalone = "true"
attrName = "drug"
createAlternateTaxTable attrStandalone, attrName


' Build <tax-tables>
attrMerchantCalculated = "true"
createTaxTables attrMerchantCalculated


' Specify A URL to which Google Checkout should send Merchant Calculations API
' (<merchant-calculation-callback>) requests and create the
' <merchant-calculations> XML for a Checkout API request.
'
' +++ CHANGE ME +++
' If you are implementing the Merchant Calculations API, you need to
' uncomment the following lines of code, which create the
' <merchant-calculations> XML in a Checkout API response. You also
' need to update the value of the $merchant_calculations_url variable
' to the URL to which Google Checkout should send 
' <merchant-calculation-callback> requests.
elemMerchantCalculationsUrl = _
    "http://www.example.com/shopping/MerchantCalculationCallback.asp"
elemAcceptMerchantCoupons = "true"
elemAcceptGiftCertificates = "true"
createMerchantCalculations elemMerchantCalculationsUrl, _
    elemAcceptMerchantCoupons, elemAcceptGiftCertificates

' +++ CHANGE ME +++
' The $edit_cart_url variable identifies a URL that the customer can
' link to to edit the contents of the shopping cart.
' The $continue_shopping_url variable identifies a URL that the
' customer can link to to continue shopping.
' If you are providing of these options to your customers, you need
' to insert the appropriate URLs for these variables.
' e.g. $edit_cart_url = "http://www.example.com/shopping/edit";
'      $continue_shopping_url = "http://www.example.com/shop/continue";
elemEditCartUrl = "http://www.example.com/shopping/edit"
elemContinueShoppingUrl = "http://www.example.com/shopping/continue"

' Build the <merchant-checkout-flow-support> element in the Checkout
' API request.
createMerchantCheckoutFlowSupport elemEditCartUrl, elemContinueShoppingUrl


' Create the shopping cart XML and HMAC-SHA1 signature that
' will be included in your HTTP POST request to Google Checkout 
'   1. Retrieve the shopping cart XML
'   2. Base64-encode the shopping cart XML and the signature;
'      a form on your web page will contain the encoded values 

Dim xmlCart
Dim b64signature
Dim b64cart

' 1. Get <checkout-shopping-cart> XML
xmlCart = createCheckoutShoppingCart

' 2. Calculate the HMAC-SHA1 value and Base64-encode the 
' Cart XML and the signature before posting
b64cart = cryptObj.base64Encode(xmlCart)
b64signature = cryptObj.generateSignature(xmlCart, strMerchantKey)

checkoutPostData = "cart=" & Server.urlencode(b64cart) & _
    "&signature=" & Server.urlencode(b64signature)

' Log <checkout-shopping-cart> XML
logMessage logFilename, checkoutPostData

' The following HTML page displays some information about the POST
' request that will be submitted to Google Checkout if you click the 
' Google Checkout button that appears on the page. The Google Checkout 
' button is embedded in a form similar to the form you want to include 
' on your site. The form sends the request to Google Checkout and shows 
' you an interface similar to what your customer would see after clicking 
' the Google Checkout button.
'
' Note: This page also calls the displayDiagnoseResponse function,
' which is defined in GlobalAPIFunctions.asp, to verify that the 
' API request contains valid XML. If the request does not contain
' valid XML, you will see a link to a tool that lets you edit and
' recheck the XML. The code for that tool is in the 
' <b>DebuggingTool.asp</b> file, which is also included
' in the <b>checkout-asp-samplecode.zip</b> file.
%>
<html>
<head>
    <style type="text/css">@import url(googleCheckout.css);</style>
</head>
<body>
    <p style="text-align:center">
    <table class="table-1" cellspacing="5" cellpadding="5">
        <tr><td style="padding-bottom:20px;text-align:center"><h2>
        Place a New Order
        </h2></td></tr>

        <!-- Print the shopping cart XML -->
        <tr><td style="padding-bottom:20px">
            <p><b>This is the &lt;checkout-shopping-cart&gt; XML before it is 
            base64-encoded:</b></p>
            <p><%=Server.HTMLEncode(xmlCart)%></p>
        </td></tr>

        <!-- Print the HMAC-SHA1 signature -->
        <tr><td style="padding-bottom:20px">
            <p><b>This is the base64-encoded signature:</b></p>
            <p><%=Server.HTMLEncode(b64signature)%></p>
        </td></tr>

        <!-- Print Error message if the cart XML is invalid -->
<%
        displayDiagnoseResponse checkoutPostData, checkoutDiagnoseUrl, _
            xmlCart, "debug"
%>
        <tr><td style="padding-bottom:20px">
            <p><b>Click on the button to post this cart.</b></p>
            <p><form method="POST" action="<%=checkoutUrl%>">
                <input type="hidden" name="cart" value="<%=b64cart%>">
                <input type="hidden" name="signature" value="<%=b64signature%>">
                <input type="image" name="Checkout" alt="Checkout" 
                src="http://checkout.google.com/buttons/checkout.gif?
                merchant_id=<%=strMerchantId%>&
                w=180&h=46&style=white&variant=text&loc=en_US" 
                height="46" width="180">
                </form></p>
        </td></tr>
    </table>
    </p>
</body>
</html>
<%
    ' Free object
    Set cryptObj = Nothing
    Set domTaxArea = Nothing
    Set domShippingRestrictions = Nothing
%>