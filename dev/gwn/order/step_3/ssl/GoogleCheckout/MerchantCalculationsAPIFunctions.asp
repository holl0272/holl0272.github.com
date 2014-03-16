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
' "MerchantCalculationsAPIFunctions.asp" contains a set of functions that
' process <merchant-calculation-callback> message 
' and build <merchant_calculation-result> message.
'*******************************************************************************


'*******************************************************************************
' The processMerchantCalculationCallback function handles a
' <merchant-calculation-callback> request and returns a
' <merchant-calculation-results> XML response. This function calls
' the createMerchantCalculationResults function, which constructs
' the <merchant-calculation-results> response. This function then
' prints the <merchant-calculation-results> response to return the
' <merchant-calculation-results> information to Google Checkout and 
' logs the response as well.
'
' Input:  domMcCallbackObj  <merchant-calculation-callback> XML
'*******************************************************************************
Function processMerchantCalculationCallback(domMcCallbackObj)

    Dim xmlMcResults

    ' Process <merchant-calculation-callback> and create 
    ' <merchant-calculation-results>
    xmlMcResults = createMerchantCalculationResults(domMcCallbackObj)

    ' Respond with <merchant-calculation-results> XML
    Response.write xmlMcResults

    ' Log <merchant-calculation-results>
    logMessage logFilename, xmlMcResults

End Function


'**************************************************************************
' The createMerchantCalculationResults function creates the XML DOM for
' a <merchant-calculation-results> XML response. This function receives
' the <merchant-calculation-callback> from the
' processMerchantCalculationCallback function.
'
' This function calls the createMerchantCodeResults, getShippingRate
' and getTaxRate functions to calculate shipping costs, taxes and
' discounts that should be applied to the order total
'
' Input:    domMcCallbackObj    <merchant-calculation-callback> XML
' Returns:  <merchant-calculation-results> XML
'**************************************************************************
Function createMerchantCalculationResults(domMcCallbackObj)

    ' Define the objects used to create the <merchant-calculation-callback>
    Dim domMcResultsObj
    Dim domMcResults
    Dim domMerchantCodeResults
    Dim domMerchantCodeResultsRoot
    Dim domResults
    Dim domResult
    Dim domResponse
    Dim domMcCallbackObjRoot
    Dim domTaxList
    Dim calcTax
    Dim domMethodList
    Dim domMethod
    Dim attrShippingName
    Dim domAnonymousAddressList
    Dim domAnonymousAddress
    Dim attrAddressId
    Dim domMerchantCodeList
    Dim totalTax
    Dim domTotalTax
    Dim domShippingRate
    Dim shippingRate
    Dim domShippable
    Dim shippable

    Set domMcResultsObj = Server.CreateObject(strMsxmlDomDocument)
    domMcResultsObj.async = False
    domMcResultsObj.appendChild( _
        domMcResultsObj.createProcessingInstruction("xml", _
            strXmlVersionEncoding))

    ' Create root tag for <merchant-calculation-results> response
    ' and set xmlns attribute
    Set domMcResults = domMcResultsObj.appendChild( _
        domMcResultsObj.createElement("merchant-calculation-results"))
    domMcResults.setAttribute "xmlns", strXmlns

    ' Create child element <results>
    Set domResults = domMcResults.appendChild( _
        domMcResultsObj.createElement("results"))

    Set domMcCallbackObjRoot = domMcCallbackObj.documentElement

    ' Retrieve Boolean value indicating whether merchant calculates
    ' tax for the order.
    '     e.g. <tax>true</tax>
    ' If you do not use custom calculations to calculate tax, you
    ' may ignore the next two lines of code.
    Set domTaxList = domMcCallbackObjRoot.getElementsByTagname("tax")
    calcTax = domTaxList(0).text

    ' Retrieve the names of the shipping methods available for the order
    ' These shipping methods will have been communicated to Google Checkout
    ' in a CheckoutAPIRequest. Note: The <merchant-calculated-callback>
    ' will only contain <merchant-calculated-shipping> options from
    ' the Checkout API request.
    Set domMethodList = _
        domMcCallbackObjRoot.getElementsByTagname("method")

    ' Retrieve shipping addresses from the <merchant-calculated-callback>
    ' response. These shipping addresses are anonymous, meaning they
    ' only include the city, region (state), postal code and country
    ' code for the address
    Set domAnonymousAddressList =  _
        domMcCallbackObjRoot.getElementsByTagname( _
            "anonymous-address")

    ' Retrieve a list of coupon and gift certificate codes that
    ' should be applied to the order total. Note: The
    ' <merchant-calculated-callback> can only contain these codes if
    ' the <accept-merchant-coupons> or <accept-gift-certificates> tag
    ' in the corresponding Checkout API request has a value of "true".
    Set domMerchantCodeList = _
        domMcCallbackObjRoot.getElementsByTagname( _
            "merchant-code-string")
    
    ' Loop through address IDs to build <result> elements
    For Each domAnonymousAddress In domAnonymousAddressList

        ' Retrieve the address ID
        attrAddressId = domAnonymousAddress.getAttribute("id")
        
        If domMethodList.length > 0 Then

            ' Loop for each merchant-calulated shipping method
            For Each domMethod In domMethodList

                ' Retrieve the name of the shipping method
                attrShippingName = domMethod.getAttribute("name")

                ' Create a <result> element in the response with
                ' shipping-name and address-id attributes
                Set domResult = _
                    domResults.appendChild( _
                        domMcResultsObj.createElement("result"))
                domResult.setAttribute "shipping-name", attrShippingName
                domResult.setAttribute "address-id", attrAddressId

                ' If the <tax> tag in the <merchant-calculation-callback>
                ' has a value of "true", call the getTaxRate function
                ' to calculate taxes for the order.
                If calcTax = "true" Then
                    Set domTotalTax = _
                        domResult.appendChild( _
                            domMcResultsObj.createElement("total-tax"))
                    domTotalTax.setAttribute "currency", attrCurrency
                    totalTax = getTaxRate(domMcCallbackObj, _
                        attrAddressId, attrShippingName)
                    domTotalTax.appendChild( _
                        domMcResultsObj.createTextNode(totalTax))
                End If

                ' If there are coupon or gift certificate codes, call
                ' the createMerchantCodeResults function to verify those
                ' codes and to create <coupon-result> or
                ' <gift-certificate-result> elements to be included in
                ' the <merchant-calculation-response>.
                If domMerchantCodeList.length > 0 Then

                    Set domMerchantCodeResults = _
                        createMerchantCodeResults(domMcCallbackObj, _
                            domMerchantCodeList, attrAddressId)

                    Set domMerchantCodeResultsRoot = _
                        domMerchantCodeResults.documentElement

                    domResult.appendChild( _
                        domMerchantCodeResultsRoot.cloneNode(true))

                End If

                ' Call the getShippingRate function to calculate the
                ' shipping cost for the shipping method-address ID
                ' combination.
                Set domShippingRate = _
                    domResult.appendChild( _
                        domMcResultsObj.createElement("shipping-rate"))
                domShippingRate.setAttribute "currency", attrCurrency
                shippingRate = getShippingRate(domMcCallbackObj, _
                    attrAddressId, attrShippingName)
                domShippingRate.appendChild( _
                    domMcResultsObj.createTextNode(shippingRate))

                ' Verify that the order can be shipped to the address
                shippable = verifyShippable(domMcCallbackObj, _
                    attrAddressId, attrShippingName)
                Set domShippable = _
                    domResult.appendChild( _
                        domMcResultsObj.createElement("shippable"))
                domShippable.text = shippable
            Next

        ' This block executes if no shipping methods are specified
        Else
            ' Create a <result> element in the response with
            ' shipping-name and address-id attributes

            Set domResult = domResults.appendChild( _
                domMcResultsObj.createElement("result"))
            domResult.setAttribute "address-id", attrAddressId

            ' If the <tax> tag in the <merchant-calculation-callback>
            ' has a value of "true", call the getTaxRate function
            ' to calculate taxes for the order.
            If calcTax = "true" Then
                Set domTotalTax = _
                    domResult.appendChild( _
                        domMcResultsObj.createElement("total-tax"))
                domTotalTax.setAttribute "currency", attrCurrency
                totalTax = getTaxRate(domMcCallbackObj, _
                    attrAddressId, attrShippingName)
                domTotalTax.appendChild( _
                    domMcResultsObj.createTextNode(totalTax))
            End If

            ' If there are coupon or gift certificate codes, call
            ' the createMerchantCodeResults function to verify those
            ' codes and to create <coupon-result> or
            ' <gift-certificate-result> elements to be included in
            ' the <merchant-calculation-response>.
            If domMerchantCodeList.length > 0 Then

                Set domMerchantCodeResults = _
                    createMerchantCodeResults(domMcCallbackObj, _
                        domMerchantCodeList, attrAddressId)

                Set domMerchantCodeResultsRoot = _
                    domMerchantCodeResults.documentElement

                domResult.appendChild( _
                    domMerchantCodeResultsRoot.cloneNode(true))

            End If
        End If
    Next

    ' Return <merchant-calculation-results> XMLDOM
    createMerchantCalculationResults = domMcResults.xml

    Set domMcResultsObj = Nothing
    Set domMcResults = Nothing
    Set domMerchantCodeResults = Nothing
    Set domMerchantCodeResultsRoot = Nothing
    Set domResults = Nothing
    Set domResult = Nothing
    Set domResponse = Nothing
    Set domMcCallbackObjRoot = Nothing
    Set domTotalTax = Nothing
    Set domShippingRate = Nothing
    Set domShippable = Nothing
    Set domTaxList = Nothing
    Set domMethodList = Nothing
    Set domMethod = Nothing
    Set domAnonymousAddressList = Nothing
    Set domAnonymousAddress = Nothing
    Set domMerchantCodeList = Nothing

End Function


'**************************************************************************
' The createMerchantCodeResults function creates the XML DOM for a
' <coupon-result> or a <gift-certificate-result> for a Merchant
' Calculations API response. This function calls the
' getMerchantCodeInfo function, which you will need to modify,
' to retrieve information about each coupon or gift certificate code.
'
' Input:    domMcCallbackObj        <merchant-calculation-callback> XML
' Input:    domMerchantCodeList    array of merchant-code-string codes
' Input:    addressId          
' Returns:  <merchant-code-results> XMLDOM
'**************************************************************************
Function createMerchantCodeResults(domMcCallbackObj, _
    domMerchantCodeList, addressId)

    ' Define the objects used to create the <coupon-result> or
    ' <gift-certificate-result>
    Dim code
    Dim domMcResultsObj
    Dim merchantCode
    Dim codeType
    Dim calculatedAmount
    Dim message
    Dim domMerchantCodeResults
    Dim domMerchantCodeResultObj
    Dim domMerchantCodeResultRoot

    ' Create an empty XMLDOM
    Set domMcResultsObj = Server.CreateObject(strMsxmlDomDocument)
    domMcResultsObj.async = False
    Set domMerchantCodeResults = _
        domMcResultsObj.appendChild( _
            domMcResultsObj.createElement("merchant-code-results"))
    
    For Each merchantCode In domMerchantCodeList

        code = merchantCode.getAttribute("code")

        Set domMerchantCodeResultObj = _
            getMerchantCodeInfo(domMcCallbackObj, code, addressId)

        Set domMerchantCodeResultRoot = _
            domMerchantCodeResultObj.documentElement

        domMerchantCodeResults.appendChild( _
            domMerchantCodeResultRoot.cloneNode(true))
        
    Next

    Set createMerchantCodeResults = domMcResultsObj    

    ' Release the objects used to create the <coupon-result> or
    ' <gift-certificate-result>
    Set domMcResultsObj = Nothing
    Set domMerchantCodeResultObj = Nothing
    Set domMerchantCodeResultRoot = Nothing

End Function


'**************************************************************************
' The getMerchantCodeInfo function retrieves information about a coupon
' or gift certificate code provided by the customer. You will need to
' modify this function to retrieve information about the code. The
' changes you will need to make are discussed in the comments in the
' function. After retrieving this information, this function calls and
' returns the value of the createMerchantCodeResult function.
'
' Input:    domMcCallbackObj    <merchant-calculation-callback> XMLDOM
' Input:    elemCode            A coupon or gift certificate code.
' Input:    addressId           An ID the corresponds to the address
'                                   to which an order should be shipped.
' Returns:  merchant-calculated shipping rate
'**************************************************************************
Function getMerchantCodeInfo(domMcCallbackObj, elemCode, addressId)

    ' Define objects that contain information about the merchant code
    Dim elemCodeType
    Dim elemCodeValid
    Dim elemCalculatedAmount
    Dim elemMessage

    ' +++ CHANGE ME +++
    ' You need to modify this function to retrieve information about
    ' a coupon or gift certificate code provided by the customer. This
    ' function needs to retrieve the following information about the code:
    '     1. The code's type. The code type may be either "coupon" or
    '         "gift-certificate".
    '     2. A flag that indicates whether the code is valid. The value
    '         of this flag must be either "true" or "false".
    '     3. The calculated amount of the code. If the code is valid,
    '         you need to quantify the amount of the code discount.
    '         This data is optional.
    '     4. A message that should be displayed with the code. This
    '         data is optional.
    ' This function returns the result from the createMerchantCodeResult
    ' function, which is a <coupon-result> or a <gift-certificate-result>,
    ' to the createMerchantCodeResults function, which adds the XML
    ' block to the response.

    elemCodeType = "coupon"
    elemCodeValid = "true"
    elemCalculatedAmount = "10.00"
    elemMessage = "You saved $" & elemCalculatedAmount

    Set getMerchantCodeInfo = _
        createMerchantCodeResult(elemCodeType, elemCodeValid, elemCode, _
            elemCalculatedAmount, elemMessage)

End Function


'**************************************************************************
' The createMerchantCodeResult function creates the XML DOM for a
' <coupon-result> or <gift-certificate-result> in a Merchant
' Calculations API response.
'
' Input:       elemCodeType           The type of code provided by the
'                                     customer. Valid values are "coupon"
'                                     and "gift-certificate".
' Input:       elemCodeValid          Indicates whether the code is valid.
'                                     Valid values are "true" and "false".
' Input:       elemCode               The code entered by the user
' Input:       elemCalculatedAmount   The amount to deduct from the total
' Input:       elemMessage            A message to display in regard to
'                                     the code.
' Returns:     <coupon-result> or <gift-certificate-result> XMLDOM
'**************************************************************************
Function createMerchantCodeResult(elemCodeType, elemCodeValid, elemCode, _
    elemCalculatedAmount, elemMessage)

    ' Define objects used to create the <coupon-result> or
    ' <gift-certificate-result>
    Dim domCodeResultObj
    Dim domMerchantCodeResult
    Dim domValid
    Dim domCode
    Dim domMessage
    Dim domCalculatedAmount

    ' create an empty XMLDOM
    Set domCodeResultObj = Server.CreateObject(strMsxmlDomDocument)
    domCodeResultObj.async = False

    ' Create root tag for <coupon-result> or <gift-certificate-result>
    Set domMerchantCodeResult = _
        domCodeResultObj.appendChild( _
            domCodeResultObj.createElement(elemCodeType & "-result"))

    ' Create <valid> tag, which will indicate whether the code is valid
    Set domValid = _
        domMerchantCodeResult.appendChild( _
            domCodeResultObj.createElement("valid"))
    domValid.text = elemCodeValid
    
    ' Add the coupon or gift certificate code in a <code> tag
    Set domCode = _
        domMerchantCodeResult.appendChild( _
            domCodeResultObj.createElement("code"))
    domCode.text = elemCode

    ' Add the <calculated-amount> tag if there is a value for the
    ' elemCalculatedAmount parameter. You could omit this tag if the
    ' code is invalid.
    If elemCalculatedAmount <> "" Then
        Set domCalculatedAmount = domMerchantCodeResult.appendChild( _
            domCodeResultObj.createElement("calculated-amount"))
        domCalculatedAmount.setAttribute "currency", attrCurrency
        domCalculatedAmount.text = elemCalculatedAmount
    End If

    ' Add a <message> tag if the $message parameter has a value
    If elemMessage <> "" Then
        Set domMessage = _
            domMerchantCodeResult.appendChild( _
                domCodeResultObj.createElement("message"))
        domMessage.text = elemMessage
    End If

    Set createMerchantCodeResult = domCodeResultObj

    ' Release objects used to create the <coupon-result> or
    ' <gift-certificate-result>
    Set domCodeResultObj = Nothing
    Set domMerchantCodeResult = Nothing
    Set domValid = Nothing
    Set domCode = Nothing
    Set domCalculatedAmount = Nothing
    Set domMessage = Nothing

End Function


'*******************************************************************************
' The verifyShippable function determines whether an order can be
' shipped to the specified address using the specified shipping method.
' You will need to modify this function to return a Boolean value
' indicating whether the order is shippable using the given shipping method.
'
' Input:    domMcCallbackObj    <merchant-calculation-callback> XMLDOM
' Input:    addressId           An ID the corresponds to the address 
'                                   to which an order should be shipped.
' Input:    shippingMethod      A shipping option for an order
' Returns:  Boolean value indicating whether items can be shipped to
'           specified address
'*******************************************************************************
Function verifyShippable(domMcCallbackObj, addressId, shippingMethod)
    ' +++ CHANGE ME +++
    ' You need to modify this function to return a Boolean (true/false)
    ' value that indicates whether the order can be shipped to the
    ' specified address (addressId) using the specified shipping
    ' method (shippingMethod).
    verifyShippable = "true"
End Function


'*******************************************************************************
' The getShippingRate function determines the cost of shipping
' the order to the specified address using the specified shipping method.
' You will need to modify this function to calculate and return this cost.
'
' Input:    domMcCallbackObj    <merchant-calculation-callback> XMLDOM
' Input:    addressId           An ID the corresponds to the address 
'                                   to which an order should be shipped.
' Input:    shippingMethod      A shipping option for an order
' Returns:  merchant-calculated shipping rate
'*******************************************************************************
Function getShippingRate(domMcCallbackObj, addressId, shippingMethod)

    ' +++ CHANGE ME +++
    ' You need to modify this function to return the cost of
    ' shipping an order to the specified address using the specified
    ' shipping method.
    getShippingRate = "8.76"

End Function


'*******************************************************************************
' The getTaxRate function returns the total tax that should be applied to
' the order if it is shipped to the specified address. You will need to
' modify this function to return the calculated tax amount.
'
' Input:    domMcCallbackObj    <merchant-calculation-callback> XMLDOM
' Input:    addressId           An ID the corresponds to the address 
'                                   to which an order should be shipped.
' Returns:  merchant-calculated total tax
'*******************************************************************************
Function getTaxRate(domMcCallbackObj, addressId, shippingMethod)

    ' +++ CHANGE ME +++
    ' You need to modify this function to return the total tax for
    ' an order based on the specified address.
    getTaxRate = "17.55"

End Function

%>