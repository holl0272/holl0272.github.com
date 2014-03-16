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
' "CheckoutAPIFunctions.asp" is a client library of functions that enable
' merchants to systematically generate Cart XML.
' 
' You should also look at "CheckoutShoppingCartDemo.asp" file to learn more 
' about how to call each function and what it returns.
'*******************************************************************************

Dim domItemsObj
Dim domShoppingCartObj
Dim domDefaultTaxRulesObj
Dim domAltTaxRulesObj
Dim domAltTaxTablesObj
Dim domTaxTablesObj
Dim domShippingRestrictionsObj
Dim domShippingMethodsObj
Dim domMerchantCalculationsObj 
Dim domMerchantCFSObj
Dim domCFSObj 
Dim domCheckoutShoppingCartObj


'*******************************************************************************
' The createItem function constructs the XML for a single
' <item> in a shopping cart.
' 
' Input:       elemItemName                 item name
' Input:       elemItemDescription          item description
' Input:       elemQuantity                 quantity
' Input:       elemUnitPrice                unit price
' Input:       elemTaxTableSelector         name of the tax table to select
'                                           for this item
' Input:       elemMerchantPrivateItemData  XML to be appended as a child to
'                                           <merchant-private-item-data>
'
' Returns:     <item> XML DOM
'*******************************************************************************
Function createItem(elemItemName, elemItemDescription, elemQuantity, _
    elemUnitPrice, elemTaxTableSelector, elemMerchantPrivateItemData)

    Dim strFunctionName
    Dim errorType
    strFunctionName = "createItem()"

    ' Each of these parameters must have a value to create an <item>
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemItemName", elemItemName
    checkForError errorType, strFunctionName, "elemItemDescription", _
        elemItemDescription
    checkForError errorType, strFunctionName, "elemQuantity", elemQuantity
    checkForError errorType, strFunctionName, "elemUnitPrice", elemUnitPrice
    checkForError errorType, strFunctionName, "attrCurrency", attrCurrency

    ' HTML entities need to be escaped properly
    elemItemName = Server.HTMLEncode(elemItemName)
    elemItemDescription = Server.HTMLEncode(elemItemDescription)

    ' Define objects used to create the item
    Dim domItemObj
    Dim domItem
    Dim domItemName
    Dim domItemDescription
    Dim domQuantity
    Dim domUnitPrice
    Dim domTaxTableSelector
    Dim domMerchantPrivateItemDataObj
    Dim domNewMerchantPrivateItemData
    Dim domMerchantPrivateItemDataRoot
    Dim domItemsRoot
    Dim domItemRoot

    ' Create the <items> tag if this is the first item to be created
    If Not(IsObject(domItemsObj)) Then
        Set domItemsObj = Server.CreateObject(strMsxmlDomDocument)
        domItemsObj.async = False
        domItemsObj.appendChild(domItemsObj.createElement("items"))
    End If

    Set domItemObj = Server.CreateObject(strMsxmlDomDocument)
    domItemObj.async = False

    ' Create the <item> tag for the item to be created
    Set domItem = domItemObj.appendChild(domItemObj.createElement("item"))

    ' Add the item name to the XML
    Set domItemName = domItem.appendChild(domItemObj.createElement("item-name"))
    domItemName.Text = elemItemName
    
    ' Add the item description to the XML
    Set domItemDescription = _
        domItem.appendChild(domItemObj.createElement("item-description"))
    domItemDescription.Text = elemItemDescription

    ' Add the quantity to the XML
    Set domQuantity = _
        domItem.appendChild(domItemObj.createElement("quantity"))
    domQuantity.Text = elemQuantity

    ' Add the unit price for the item to the XML
    Set domUnitPrice = _
        domItem.appendChild(domItemObj.createElement("unit-price"))
    domUnitPrice.setAttribute "currency", attrCurrency
    domUnitPrice.Text = elemUnitPrice

    ' If there is an alternate-tax-table associated with this item,
    ' specify the table's name using the <tax-table-selector> tag.
    If elemTaxTableSelector <> "" Then
        Set domTaxTableSelector = _
            domItem.appendChild(domItemObj.createElement("tax-table-selector"))
        domTaxTableSelector.Text = elemTaxTableSelector
    End If

    ' If you have provided a value for the elemMerchantPrivateItemData
    ' variable, that value will be printed inside the
    ' <merchant-private-item-data> tag.
    If elemMerchantPrivateItemData <> "" Then

        Set domMerchantPrivateItemDataObj = _
            Server.CreateObject(strMsxmlDomDocument)
        domMerchantPrivateItemDataObj.async = False
        domMerchantPrivateItemDataObj.loadXml elemMerchantPrivateItemData

        Set domNewMerchantPrivateItemData = domItem.appendChild( _
            domItemObj.createElement("merchant-private-item-data"))

        Set domMerchantPrivateItemDataRoot = _
            domMerchantPrivateItemDataObj.documentElement

        domNewMerchantPrivateItemData.appendChild( _
            domMerchantPrivateItemDataRoot.cloneNode(True))

    End If

    ' The newly created item is added as a child of the <items> tag.
    Set domItemsRoot = domItemsObj.documentElement
    Set domItemRoot = domItemObj.documentElement
    domItemsRoot.appendChild domItemRoot.cloneNode(True)

   Set createItem = domItemObj

    ' Release objects used to create item
    Set domItemObj = Nothing
    Set domItem = Nothing
    Set domItemName = Nothing
    Set domItemDescription = Nothing
    Set domQuantity = Nothing
    Set domUnitPrice = Nothing
    Set domTaxTableSelector = Nothing
    Set domMerchantPrivateItemDataObj = Nothing
    Set domNewMerchantPrivateItemData = Nothing
    Set domMerchantPrivateItemDataRoot = Nothing
    Set domItemsRoot = Nothing
    Set domItemRoot = Nothing

End Function


'*******************************************************************************
' The createShoppingCart function constructs the XML for the
' <shopping-cart> element in a Checkout API request. Since the
' <shopping-cart> element contains all of the items in the cart,
' this function must be called after you have already called
' the createItem function for each item in the order.
'
' Input:       dtmCartExpiration         date and time in 
'                                        "yyyy-mm-ddThh:mm:ss" format
' Input:       elemMerchantPrivateData   XML to be appended as a child to 
'                                        <merchant-private-data>
'
' Returns:     <shopping-cart> XML DOM
'*******************************************************************************
Function createShoppingCart(dtmCartExpiration, elemMerchantPrivateData)

    Dim strFunctionName
    Dim errorType

    strFunctionName = "createShoppingCart()"
    
    ' There must be at least one item in the shopping cart by the
    ' time you call this function or the function will log an error.
    errorType = "MISSING_PARAM"
    If Not(IsObject(domItemsObj)) Then
        errorHandler errorType, strFunctionName, "domItems"
    End If 

    ' Define objects used to create the shopping cart
    Dim domShoppingCart
    Dim domItemsRoot
    Dim domCartExpiration
    Dim domGoodUntilDate
    Dim domNewMerchantPrivateData
    Dim domMerchantPrivateDataObj
    Dim domMerchantPrivateDataRoot

    Set domShoppingCartObj = Server.CreateObject(strMsxmlDomDocument)
    domShoppingCartObj.async = False

    ' Create the <shopping-cart> element
    Set domShoppingCart = domShoppingCartObj.appendChild( _
        domShoppingCartObj.createElement("shopping-cart"))
    Set domItemsRoot = domItemsObj.documentElement
    domShoppingCart.appendChild(domItemsRoot.cloneNode(true))
    
    ' If there is an expiration date ($cart_expiration) for the cart,
    ' include it in the <shopping-cart> XML.
    If dtmCartExpiration <> "" Then
        Set domCartExpiration = domShoppingCart.appendChild( _
            domShoppingCartObj.createElement("cart-expiration"))
        Set domGoodUntilDate = domCartExpiration.appendChild( _
            domShoppingCartObj.createElement("good-until-date"))
        domGoodUntilDate.Text = dtmCartExpiration
    End If

    ' If you have provided a value for the $merchant_private_data
    ' variable, that value will be printed inside the
    ' <merchant-private-data> tag.
    If elemMerchantPrivateData <> "" Then

        Set domMerchantPrivateDataObj = Server.CreateObject(strMsxmlDomDocument)
        domMerchantPrivateDataObj.async = False
        domMerchantPrivateDataObj.loadXml elemMerchantPrivateData

        Set domNewMerchantPrivateData = domShoppingCart.appendChild( _
            domShoppingCartObj.createElement("merchant-private-data"))

        Set domMerchantPrivateDataRoot = _
            domMerchantPrivateDataObj.documentElement

        domNewMerchantPrivateData.appendChild( _
            domMerchantPrivateDataRoot.cloneNode(true))
    End If

    Set createShoppingCart = domShoppingCartObj

    ' Release objects used to create shipping cart
    Set domShoppingCart = Nothing
    Set domItemsObj = Nothing
    Set domItemsRoot = Nothing
    Set domCartExpiration = Nothing
    Set domGoodUntilDate = Nothing
    Set domNewMerchantPrivateData = Nothing
    Set domMerchantPrivateDataObj = Nothing
    Set domMerchantPrivateDataRoot = Nothing

End Function


'*******************************************************************************
' The createUsCountryArea function is a wrapper function that calls
' the createUsPlaceArea function. The createUsPlaceArea function, in turn,
' creates and returns a <us-country-area> XML DOM.
'
' Input:       areaPlace       The U.S. region that should be included in 
'                              the XML block. Valid values are
'                              CONTINENTAL_48, FULL_50_STATES and ALL.
'
' Returns:     <us-country-area> XML DOM
'*******************************************************************************
Function createUsCountryArea(areaPlace)
    Set createUsCountryArea = createUsPlaceArea("country", areaPlace)
End Function


'*******************************************************************************
' The createUsStateArea function is a wrapper function that calls
' the createUsPlaceArea function. The createUsPlaceArea function, in turn,
' creates and returns a <us-state-area> XML DOM.
'
' Input:       areaPlace        The U.S. state that should be included
'                               in the XML block. The value should be a
'                               two-letter U.S. state abbreviation.
'
' Returns:     <us-state-area> XML DOM
'*******************************************************************************
Function createUsStateArea(areaPlace)
    Set createUsStateArea = createUsPlaceArea("state", areaPlace)
End Function


'*******************************************************************************
' The createUsZipArea function is a wrapper function that calls the
' createUsPlaceArea function. The createUsPlaceArea function, in turn,
' creates and returns a <us-zip-area> XML block.
'
' Input:       areaPlace       The zip code that should be included
'                              in the XML block. The value should be a
'                              five-digit zip code or a zip code pattern.
'
' Returns:     <us-zip-area> XML DOM
'*******************************************************************************
Function createUsZipArea(areaPlace)
    Set createUsZipArea = createUsPlaceArea("zip", areaPlace)
End Function


'*******************************************************************************
' The createUsPlaceArea function creates <us-country-area>,
' <us-state-area> and <us-zip-area> XML blocks.
'
' Input:       areaType        The type of XML object to be created. Valid
'                              values are "country", "state" and "zip".
' Input:       areaPlace       This value corresponds to the accepted
'                              areaPlace parameter values for the
'                              createUsCountryArea, createUsStateArea and
'                              createUsZipArea functions.
'
' Returns:     <us-country-area>, <us-state-area> or <us-zip-area> XML DOM
'*******************************************************************************
Function createUsPlaceArea(areaType, areaPlace)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createUsPlaceArea()"
    
    ' Both parameters must be specified for the function call to execute.
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "areaType", areaType
    checkForError errorType, strFunctionName, "areaPlace", areaPlace

    ' Define objects used to create the XML block
    Dim domAreaObj
    Dim domArea
    Dim domAreaPlace

    ' Create the parent XML element for the areaType
    Set domAreaObj = Server.CreateObject(strMsxmlDomDocument)
    domAreaObj.async = False
    Set domArea = domAreaObj.appendChild( _
        domAreaObj.createElement("us-" & areaType & "-area"))

    ' Create the element that contains the areaPlace data
    If areaType = "state" Then

        Set domAreaPlace = _
            domArea.appendChild(domAreaObj.createElement("state"))
        domAreaPlace.Text = areaPlace

    ElseIf areaType = "zip" Then

        Set domAreaPlace = _
            domArea.appendChild(domAreaObj.createElement("zip-pattern"))
        domAreaPlace.Text = areaPlace

    ElseIf areaType = "country" Then

        domArea.setAttribute "country-area", areaPlace

    End If

    Set createUsPlaceArea =  domAreaObj

    ' Release objects used to create the XML block
    Set domAreaObj = Nothing
    Set domArea = Nothing
    Set domAreaPlace = Nothing

End Function


'*******************************************************************************
' The createTaxArea function creates a <tax-area> XML DOM, which identifies
' a geographic region where a tax rate applies.
'
' Input:       taxAreaType      Valid values are "country",
'                               "state" and "zip"
' Input:       taxAreaPlace     See the valid values for the
'                               $area_place parameter of the
'                               createUsPlaceArea function
'
' Returns:     <tax-area>       <tax-area> XML DOM containing 
'                               the child elements that correspond
'                               to the specified areaType
'*******************************************************************************
Function createTaxArea(taxAreaType, taxAreaPlace) 
    
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createTaxArea()"
    
    ' Both parameters must be specified for the function call to execute.
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "taxAreaType", taxAreaType
    checkForError errorType, strFunctionName, "taxAreaPlace", taxAreaPlace

    ' Define the objects used to create the tax area
    Dim domTaxAreaObj
    Dim domTaxArea
    Dim domArea
    Dim domAreaRoot

    ' Create the <tax-area> element
    Set domTaxAreaObj = Server.CreateObject(strMsxmlDomDocument)
    domTaxAreaObj.async = False
    Set domTaxArea = _
        domTaxAreaObj.appendChild(domTaxAreaObj.createElement("tax-area"))

    ' Call the createUsPlaceArea function to create the child
    ' elements of the <tax-area> element
    Set domArea = createUsPlaceArea(taxAreaType, taxAreaPlace)
    Set domAreaRoot = domArea.documentElement
    domTaxArea.appendChild(domAreaRoot.cloneNode(true))

    Set createTaxArea = domTaxAreaObj

    ' Release the objects used to create the tax area
    Set domTaxAreaObj = Nothing
    Set domTaxArea = Nothing
    Set domArea = Nothing
    Set domAreaRoot = Nothing

End Function


'*******************************************************************************
' The createDefaultTaxRule function creates and returns a
' <default-tax-rule> XML DOM.
'
' Input:       elemRate                The tax rate to assess for a
'                                      given tax rule.
' Input:       domTaxArea              An XML DOM that identifies the
'                                      area where a tax rate should be
'                                      applied.
' Input:       elemShippingTaxed       A Boolean value that indicates
'                                      whether shipping costs are taxed
'                                      in the specified tax area.
'
' Returns:     <default-tax-rule> XML DOM
'*******************************************************************************
Function createDefaultTaxRule(elemRate, domTaxArea, elemShippingTaxed) 

    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createDefaultTaxRule()"
    
    ' Check for missing parameters
    ' You must specify a rate and provide a domTaxArea object for each rule
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemRate", elemRate

    If Not(IsObject(domTaxArea)) Then
        errorHandler errorType, strFunctionName, "domTaxArea"
    End If
    
    ' Define the objects used to create the <default-tax-rule>
    Dim domDefaultTaxRuleObj
    Dim domDefaultTaxRule
    Dim domShippingTaxed
    Dim domRate
    Dim domTaxAreaRoot
    Dim domDefaultTaxRulesRoot
    Dim domDefaultTaxRuleRoot

    ' Create the <default-tax-rule> element
    Set domDefaultTaxRuleObj = Server.CreateObject(strMsxmlDomDocument)
    domDefaultTaxRuleObj.async = False

    Set domDefaultTaxRule = domDefaultTaxRuleObj.appendChild( _
        domDefaultTaxRuleObj.createElement("default-tax-rule"))

    ' Add a <shipping-taxed> element if a elemShippingTaxed value is provided
    Set domShippingTaxed = domDefaultTaxRule.appendChild( _
        domDefaultTaxRuleObj.createElement("shipping-taxed"))

    domShippingTaxed.appendChild( _
        domDefaultTaxRuleObj.createTextNode(elemShippingTaxed))

    ' Add the tax rate for the tax rule
    Set domRate = domDefaultTaxRule.appendChild( _
        domDefaultTaxRuleObj.createElement("rate"))
    domRate.appendChild(domDefaultTaxRuleObj.createTextNode(elemRate))

    Set domTaxAreaRoot = domTaxArea.documentElement
    domDefaultTaxRule.appendChild(domTaxAreaRoot.cloneNode(true))

    ' Create a <tax-rules> element if no other <default-tax-rule>
    ' elements have been created. Append the rule to a list that
    ' will appear under the <tax-rules> element within the
    ' <default-tax-table> element
    If Not(IsObject(domDefaultTaxRulesObj)) Then
        Set domDefaultTaxRulesObj = Server.CreateObject(strMsxmlDomDocument)
        domDefaultTaxRulesObj.async = False
        domDefaultTaxRulesObj.appendChild( _
            domDefaultTaxRulesObj.createElement("tax-rules"))
    End If

    ' Add the tax rules to the XML
    Set domDefaultTaxRulesRoot = domDefaultTaxRulesObj.documentElement
    Set domDefaultTaxRuleRoot = domDefaultTaxRuleObj.documentElement
    domDefaultTaxRulesRoot.appendChild(domDefaultTaxRuleRoot.cloneNode(true))

    Set createDefaultTaxRule = domDefaultTaxRuleObj

    ' Release the objects used to create the <default-tax-rule>
    Set domDefaultTaxRuleObj = Nothing
    Set domDefaultTaxRule = Nothing
    Set domShippingTaxed = Nothing
    Set domRate = Nothing
    Set domTaxAreaRoot = Nothing
    Set domDefaultTaxRulesRoot = Nothing
    Set domDefaultTaxRuleRoot = Nothing

End Function


'*******************************************************************************
' The createAlternateTaxRule function creates and returns an
' <alternate-tax-rule> XML DOM.
'
' Input:       elemRate                The tax rate to assess for a
'                                      given tax rule.
' Input:       domTaxArea              An XML DOM that identifies the
'                                      area where a tax rate should be
'                                      applied.
'
' Returns:     <alternate-tax-rule> XML DOM
'*******************************************************************************
Function createAlternateTaxRule(elemRate, domTaxArea) 
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createAlternateTaxRule()"
    
    ' You must specify an elemRate and domTaxArea object for each tax rule
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemRate", elemRate

    If Not(IsObject(domTaxArea)) Then
        errorHandler errorType, strFunctionName, "domTaxArea"
    End If

    ' Define the objects used to create the <alternate-tax-rule>
    Dim domAltTaxRuleObj
    Dim domAltTaxRule
    Dim domRate
    Dim domTaxAreaRoot
    Dim domAltTaxRulesRoot
    Dim domAltTaxRuleRoot

    ' Create the <alternate-tax-rule> element
    Set domAltTaxRuleObj = Server.CreateObject(strMsxmlDomDocument)
    domAltTaxRuleObj.async = False
    Set domAltTaxRule = domAltTaxRuleObj.appendChild( _
        domAltTaxRuleObj.createElement("alternate-tax-rule"))

    ' Add the tax rate for the tax rule
    Set domRate = _
        domAltTaxRule.appendChild(domAltTaxRuleObj.createElement("rate"))
    domRate.appendChild(domAltTaxRuleObj.createTextNode(elemRate))

    Set domTaxAreaRoot = domTaxArea.documentElement
    domAltTaxRule.appendChild(domTaxAreaRoot.cloneNode(true))

    ' Create an <alternate-tax-rules> element if this is the first
    ' <alternate-tax-rule> to be created. Append the rule to a list
    ' that will appear under the <alternate-tax-rules> element within
    ' an <alternate-tax-table> element
    If Not(IsObject(domAltTaxRulesObj)) Then
        Set domAltTaxRulesObj = Server.CreateObject(strMsxmlDomDocument)
        domAltTaxRulesObj.async = False
        domAltTaxRulesObj.appendChild( _
            domAltTaxRulesObj.createElement("alternate-tax-rules"))
    End If

    ' Add the alternate tax rules to the XML
    Set domAltTaxRulesRoot = domAltTaxRulesObj.documentElement
    Set domAltTaxRuleRoot = domAltTaxRuleObj.documentElement
    domAltTaxRulesRoot.appendChild(domAltTaxRuleRoot.cloneNode(true))

    Set createAlternateTaxRule = domAltTaxRuleObj

    ' Release the objects used to create the <alternate-tax-rule>
    Set domAltTaxRuleObj = Nothing
    Set domAltTaxRule = Nothing
    Set domRate = Nothing
    Set domTaxAreaRoot = Nothing
    Set domAltTaxRulesRoot = Nothing
    Set domAltTaxRuleRoot = Nothing

End Function


'*******************************************************************************
' The createAlternateTaxTable function creates and returns an
' <alternate-tax-table> XML DOM. The XML will contain any
' <alternate-tax-rule> elements that have not already been included
' in an <alternate-tax-table>.
'
' Input:       attrStandalone     A Boolean value that indicates 
'                                 how taxes should be calculated if 
'                                 there is no matching <alternate-tax-rule> 
'                                 for the customer's area.
' Input:       attrNmae           A name that is used to identify 
'                                 the tax table
'
' Returns:     <alternate-tax-table> XML DOM
'*******************************************************************************
Function createAlternateTaxTable(attrStandalone, attrName) 
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createAlternateTaxTable()"
    
    ' You must specify values for the attrStandalone and attrName parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrStandalone", attrStandalone
    checkForError errorType, strFunctionName, "attrName", attrName

    ' There must be at least one alternate tax rule to include
    ' in the <alternate-tax-table>. This tax table will include
    ' any <alternate-tax-rule> elements that were created since
    ' after the last call to the createAlternateTaxTable function.

    If Not(IsObject(domAltTaxRulesObj)) Then
        errorHandler errorType, strFunctionName, "domAlternateTaxRules"
    End If
    
    ' Define the objects used to create the <alternate-tax-rule>
    Dim domAltTaxTableObj
    Dim domAltTaxTable
    Dim domAltTaxRulesRoot
    Dim domAltTaxTablesRoot
    Dim domAltTaxTableRoot

    Set domAltTaxTableObj = Server.CreateObject(strMsxmlDomDocument)
    domAltTaxTableObj.async = False

    ' Create the <alternate-tax-table> element
    Set domAltTaxTable = _
        domAltTaxTableObj.appendChild(domAltTaxTableObj.createElement("alternate-tax-table"))
    domAltTaxTable.setAttribute "standalone", attrStandalone
    domAltTaxTable.setAttribute "name", attrName

    ' Add the <alternate-tax-rules> element as
    ' a child element of <alternate-tax-table> elements
    Set domAltTaxRulesRoot = domAltTaxRulesObj.documentElement
    domAltTaxTable.appendChild(domAltTaxRulesRoot.cloneNode(true))

    ' Create an <alternate-tax-tables> element, if one has not yet
    ' been created, to contain all <alternate-tax-table> elements
    If Not(IsObject(domAltTaxTablesObj)) Then
        Set domAltTaxTablesObj = Server.CreateObject(strMsxmlDomDocument)
        domAltTaxTablesObj.async = False
        domAltTaxTablesObj.appendChild(domAltTaxTablesObj.createElement("alternate-tax-tables"))
    End If

    ' Add the <alternate-tax-table> element as a child of
    ' the <alternate-tax-tables> element
    Set domAltTaxTablesRoot = domAltTaxTablesObj.documentElement
    Set domAltTaxTableRoot = domAltTaxTableObj.documentElement
    domAltTaxTablesRoot.appendChild(domAltTaxTableRoot.cloneNode(true))

    Set domAltTaxRulesObj = Server.CreateObject(strMsxmlDomDocument)
    domAltTaxRulesObj.async = False
    domAltTaxRulesObj.appendChild(domAltTaxRulesObj.createElement("alternate-tax-rules"))

    Set createAlternateTaxTable = domAltTaxTableObj

    ' Release the objects used to create the <alternate-tax-rule>
    Set domAltTaxTableObj = Nothing
    Set domAltTaxTable = Nothing
    Set domAltTaxRulesRoot = Nothing
    Set domAltTaxTablesRoot = Nothing
    Set domAltTaxTableRoot = Nothing

End Function


'*******************************************************************************
' The createTaxTables element constructs the <tax-tables> XML DOM.
'
' Input:       attrMerchantCalculated    A Boolean value that indicates
'                                        whether tax for the order is
'                                        calculated using a special process.
'
' Returns:     <tax-tables> XML DOM
'*******************************************************************************
Function createTaxTables(attrMerchantCalculated) 
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createTaxTables()"
    
    ' Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrMerchantCalculated", _
        attrMerchantCalculated

    ' Define the objects used to create the <tax-tables>
    Dim domTaxTables
    Dim domDefaultTaxRulesRoot
    Dim domDefaultTaxTableObj
    Dim domDefaultTaxTable
    Dim domDefaultTaxTableRoot
    Dim domAltTaxTablesRoot

    ' Create the <tax-tables> element
    Set domTaxTablesObj = Server.CreateObject(strMsxmlDomDocument)
    domTaxTablesObj.async = False
    Set domTaxTables = _
        domTaxTablesObj.appendChild(domTaxTablesObj.createElement("tax-tables"))

    ' Set the "merchant-calculated" attribute on the <tax-tables> element
    If attrMerchantCalculated <> "" Then
        domTaxTables.setAttribute "merchant-calculated", attrMerchantCalculated
    End If

    ' Create a <default-tax-table> element and append the default tax rules
    Set domDefaultTaxTableObj = Server.CreateObject(strMsxmlDomDocument)
    domDefaultTaxTableObj.async = False
    Set domDefaultTaxTable = domDefaultTaxTableObj.appendChild( _
        domDefaultTaxTableObj.createElement("default-tax-table"))
    Set domDefaultTaxRulesRoot = domDefaultTaxRulesObj.documentElement
    domDefaultTaxTable.appendChild(domDefaultTaxRulesRoot.cloneNode(true))

    ' Make the <default-tax-table> element a child of <tax-tables> element
    Set domDefaultTaxTableRoot = domDefaultTaxTableObj.documentElement
    domTaxTables.appendChild(domDefaultTaxTableRoot.cloneNode(true))

    ' Add the <alternate-tax-tables> elements as children of <tax-tables>
    If IsObject(domAltTaxTablesObj) Then
        Set domAltTaxTablesRoot = domAltTaxTablesObj.documentElement
        domTaxTables.appendChild(domAltTaxTablesRoot.cloneNode(true))
    End If

    Set createTaxTables = domTaxTablesObj

    ' Release the objects used to create the <tax-tables>
    Set domTaxTables = Nothing
    Set domDefaultTaxRulesObj = Nothing
    Set domDefaultTaxRulesRoot = Nothing
    Set domDefaultTaxTableObj = Nothing
    Set domDefaultTaxTable = Nothing
    Set domDefaultTaxTableRoot = Nothing
    Set domAltTaxRulesObj = Nothing
    Set domAltTaxTablesObj = Nothing
    Set domAltTaxTablesRoot = Nothing

End Function


'*******************************************************************************
' The addAllowedAreas function is a wrapper function that calls the
' addAreas function. The addAreas function, in turn,
' creates and returns an <allowed-areas> XML DOM.
'
' Input:       attrAllowedCountry   See the attrCountry parameter
'                                   of the addAreas function for
'                                   a list of valid values
' Input:       arrayAllowedState    See the arrayState parameter
'                                   of the addAreas function for
'                                   a list of valid values
' Input:       arrayAllowedZip      See the arrayZip parameter
'                                   of the addAreas function for
'                                   a list of valid values
'
' Returns:     <shipping-restrictions> XML DOM with the allowed area added
'*******************************************************************************
Function addAllowedAreas(attrAllowedCountry, arrayAllowedState, arrayAllowedZip)

    Set addAllowedAreas = _
        addAreas(attrAllowedCountry, arrayAllowedState, arrayAllowedZip, _
            "allowed")

End Function


'*******************************************************************************
' The addExcludedAreas function is a wrapper function that calls the
' addAreas function. The addAreas function, in turn,
' creates and returns an <excluded-areas> XML DOM.
'
' Input:       attrExcludedCountry   See the attrCountry parameter
'                                    of the addAreas function for
'                                    a list of valid values
' Input:       arrayExcludedState    See the arrayState parameter
'                                    of the addAreas function for
'                                    a list of valid values
' Input:       arrayExcludedZip      See the arrayZip parameter
'                                    of the addAreas function for
'                                    a list of valid values
'
' Returns:     <shipping-restrictions> XML DOM with the excluded area added
'*******************************************************************************
Function addExcludedAreas(attrExcludedCountry, arrayExcludedState, _
    arrayExcludedZip)

    Set addExcludedAreas = _
        addAreas(attrExcludedCountry, arrayExcludedState, arrayExcludedZip, _
            "excluded")

End Function


'*******************************************************************************
' The addAreas function creates a list of regions where shipping options
' are either available or unavailable. The first three parameters identify
' the regions where the shipping option is available or unavailable. The
' final parameter indicates whether the shipping option is available.
'
' Input:       attrCountry         An region of the country where the
'                                  shipping option is either available or
'                                  unavailable. Valid values are
'                                  CONTINENTAL_48, FULL_50_STATES and ALL.
' Input:       arrayState          An array of states where the shipping
'                                  option is either available or unavailable.
'                                  Each item in the array should be a
'                                  two-letter U.S. state abbreviation.
'                                  in the XML block. The value should be a
'                                  five-digit zip code or a zip code pattern.
' Input:       arrayZip            An array of zip codes where the shipping
'                                  option is either available or unavailable.
'                                  Each item in the array should be a
'                                  five-digit zip code or a zip code pattern.
' Input:       allowedOrExcluded   Indicates whether the shipping option
'                                  is available or unavailable in the
'                                  specified regions. Valid values are
'                                  "allowed" and "excluded".
'
' Returns:     <shipping-restrictions> XML DOM with the area added
'*******************************************************************************
Function addAreas(attrCountry, arrayState, arrayZip, allowedOrExcluded)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "addAllowedAreas()"

    ' Verify that arrayState and arrayZip parameters are actually arrays
    errorType = "INVALID_INPUT_ARRAY"
    If Not(IsArray(arrayAllowedState)) Then
        errorHandler errorType, strFunctionName, "arrayState"
    End If

    If Not(IsArray(arrayAllowedZip)) Then
        errorHandler errorType, strFunctionName, "arrayZip"
    End If

    ' Verify that at least one region has been specified
    errorType = "MISSING_PARAM_NONE"
    If attrAllowedCountryArea = "" _
        And UBound(arrayAllowedState) < 0 _
        And UBound(arrayAllowedZip) < 0 _
    Then
        errorHandler errorType, strFunctionName, "attrCountry"
    End If

    ' Define objects used to create the <allowed-areas> or
    ' <excluded-areas> element
    Dim domAreasObj
    Dim domAreas
    Dim domAreasRoot
    Dim domCountry
    Dim domCountryRoot
    Dim domState
    Dim domStateRoot
    Dim domZip
    Dim domZipRoot
    Dim domShippingRestrictionsRoot
     
    Dim iUboundState
    Dim iUboundZip
    Dim iState
    Dim iZip

    ' Create the <allowed-areas> or <excluded-areas> element
    Set domAreasObj = Server.CreateObject(strMsxmlDomDocument)
    domAreasObj.async = False
    Set domAreas = domAreasObj.appendChild( _
        domAreasObj.createElement(allowedOrExcluded & "-areas"))

    ' Add the <us-country-area> element if an attrCountry is provided
    If attrCountry <> "" Then
        Set domCountry = createUsCountryArea(attrCountry)
        Set domCountryRoot = domCountry.documentElement
        domAreas.appendChild(domCountryRoot.cloneNode(true))
    End If

    ' Add a <us-state-area> element for each item in the arrayState array
    For iState = 0 To UBound(arrayState)
        If arrayState(iState) <> "" Then
            Set domState = createUsStateArea(arrayState(iState))
            Set domStateRoot = domState.documentElement
            domAreas.appendChild(domStateRoot.cloneNode(true))
        End If
    Next

    ' Add a <us-zip-area> element for each item in the arrayZip array
    For iZip = 0 To UBound(arrayZip)
        If arrayZip(iZip) <> "" Then
            Set domZip = createUsZipArea(arrayZip(iZip))
            Set domZipRoot = domZip.documentElement
            domAreas.appendChild(domZipRoot.cloneNode(true))
        End If
    Next

    ' Create a <shipping-restrictions> parent element if one has
    ' not already been created
    If Not(IsObject(domShippingRestrictionsObj)) Then
        Set domShippingRestrictionsObj = _
            Server.CreateObject(strMsxmlDomDocument)
        domShippingRestrictionsObj.async = False
        domShippingRestrictionsObj.appendChild( _
            domShippingRestrictionsObj.createElement("shipping-restrictions"))
    End If

    ' Add the shipping restrictions to the XML
    Set domShippingRestrictionsRoot = _
        domShippingRestrictionsObj.documentElement
    Set domAreasRoot = domAreasObj.documentElement
    domShippingRestrictionsRoot.appendChild(domAreasRoot.cloneNode(true))    

    Set addAreas = domShippingRestrictionsObj

    ' Release objects used to create the <allowed-areas> or
    ' <excluded-areas> element
    Set domAreasObj = Nothing
    Set domAreas = Nothing
    Set domAreasRoot = Nothing
    Set domCountry = Nothing
    Set domCountryRoot = Nothing
    Set domState = Nothing
    Set domStateRoot = Nothing
    Set domZip = Nothing
    Set domZipRoot = Nothing
    Set domShippingRestrictionsRoot = Nothing

End Function


'*******************************************************************************
' The createFlatRateShipping function is a wrapper function that calls the
' createShipping function. The createShipping function, in turn,
' creates and returns a <flat-rate-shipping> XML DOM.
'
' Input:    attrName               A name that identifies the shipping option
' Input:    elemPrice              The cost of the shipping option
' Input:    domShippingRestrictionsObj
'                                 An XML DOM that identifies areas where the 
'                                 shipping option is available or unavailable
'
' Returns:  <flat-rate-shipping> XML DOM
'*******************************************************************************
Function createFlatRateShipping(attrName, elemPrice, domShippingRestrictionsObj)

    Set createFlatRateShipping = _
        createShipping("flat-rate-shipping", attrName, elemPrice, _
            domShippingRestrictionsObj)

End Function


'*******************************************************************************
' The createMerchantCalculatedShipping function is a wrapper function
' that calls the createShipping function. The createShipping function,
' in turn, creates and returns a <merchant-calculated-shipping> XML DOM.
'
' Input:    attrName               A name that identifies the shipping option
' Input:    elemPrice              The cost of the shipping option
' Input:    domShippingRestrictionsObj
'                                  An XML DOM that identifies areas where the 
'                                  shipping option is available or unavailable
'
' Returns:  <merchant-calculated-shipping> XML DOM
'*******************************************************************************
Function createMerchantCalculatedShipping(attrName, elemPrice, _
    domShippingRestrictionsObj)

   Set createMerchantCalculatedShipping = _
        createShipping("merchant-calculated-shipping", attrName, elemPrice, _
            domShippingRestrictionsObj)

End Function


'*******************************************************************************
' The createPickup function is a wrapper function that calls the
' createShipping function. The createShipping function, in turn,
' creates and returns a <pickup> XML DOM.
'
' Input:    attrName               A name used to identify the shipping option
' Input:    elemPrice              The cost of the shipping option
'
' Returns:  <pickup> XML DOM
'******************************************************************************
Function createPickup(attrName, elemPrice)

   Set createPickup = _
        createShipping("pickup", attrName, elemPrice, "")

End Function


'*******************************************************************************
' The createShipping function creates and returns <flat-rate-shipping>,
' <merchant-calculated-shipping> or <pickup> XML DOM objects. Each call
' to this function identifies the type of shipping option, the cost of
' the shipping option as well as a name that can be used to identify the
' shipping option. The function also accepts shipping restrictions for
' <flat-rate-shipping> and <merchant-calculated-shipping>
'
' Input:   shippingType           Identifies the type of shipping. Valid values
'                                 are "flat-rate-shipping",
'                                 "merchant-calculated-shipping" and "pickup"
' Input:   attrName               A name that identifies the shipping option
' Input:   elemPrice              The cost of the shipping option
' Input:   domShippingRestrictionsObj
'                                 An XML DOM that identifies areas where the 
'                                 shipping option is available or unavailable
'
' Returns: <flat-rate-shipping>, <merchant-calculated-shipping>,
'                                 or <pickup> XML DOM
'*******************************************************************************
Function createShipping(shippingType, attrName, elemPrice, _
    domShippingRestrictionsObj)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createFlatRateShipping()"

    ' Verify that there are values for all required parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "attrName", attrName
    checkForError errorType, strFunctionName, "elemPrice", elemPrice
    checkForError errorType, strFunctionName, "attrCurrency", attrCurrency
    
    ' Define the variables used to create the shipping information
    Dim domShippingObj
    Dim domShipping
    Dim domShippingRoot
    Dim domPrice
    Dim domShippingRestrictionsRoot
    Dim domShippingMethodsRoot

    ' Create a new parent element using the shippingType as the element name
    Set domShippingObj = Server.CreateObject(strMsxmlDomDocument)
    domShippingObj.async = False
    Set domShipping = _
        domShippingObj.appendChild(domShippingObj.createElement(shippingType))

    ' Set the name and price for the shipping option
    domShipping.setAttribute "name", attrName
    Set domPrice = _
        domShipping.appendChild(domShippingObj.createElement("price"))
    domPrice.setAttribute "currency", attrCurrency
    domPrice.Text = elemPrice

    ' Add shipping-restrictions for <flat-rate-shipping> and
    ' <merchant-calculated-shipping>
    If (shippingType = "flat-rate-shipping" _
        Or shippingType = "merchant-calculated-shipping") _
        And IsObject(domShippingRestrictionsObj) _
    Then
        Set domShippingRestrictionsRoot = _
            domShippingRestrictionsObj.documentElement
        domShipping.appendChild(domShippingRestrictionsRoot.cloneNode(true))
    End If

    ' Create a <shipping-methods> element if one has not already been created
    If Not(IsObject(domShippingMethodsObj)) Then
        Set domShippingMethodsObj = Server.CreateObject(strMsxmlDomDocument)
        domShippingMethodsObj.async = False
        domShippingMethodsObj.appendChild( _
            domShippingMethodsObj.createElement("shipping-methods"))
    End If

    ' Add the shipping method to the XML request
    Set domShippingMethodsRoot = domShippingMethodsObj.documentElement
    Set domShippingRoot = domShippingObj.documentElement
    domShippingMethodsRoot.appendChild(domShippingRoot.cloneNode(true))

    Set createShipping = domShippingObj

    ' Release the variables used to create the shipping information
    Set domShippingObj = Nothing
    Set domShipping = Nothing
    Set domShippingRoot = Nothing
    Set domPrice = Nothing
    Set domShippingRestrictionsRoot = Nothing
    Set domShippingMethodsRoot = Nothing

End Function


'*******************************************************************************
' The createMerchantCalculations function creates and returns a
' <merchant-calculations> XML DOM.
'
' Input:   elemMerchantCalculationsURL   Callback URL for merchant calculations
' Input:   elemAcceptMerchantCoupons     Boolean value that indicates
'                                        whether Google Checkout should display
'                                        an option for customers to enter
'                                        coupon codes for an order
' Input:   elemAcceptGiftCertificates    Boolean value that indicates
'                                        whether Google Checkout should display
'                                        an option for customers to enter
'                                        gift certificate codes
'
' Returns:  <merchant-calculations> XML DOM
'*******************************************************************************
Function createMerchantCalculations(elemMerchantCalculationsUrl, _
    elemAcceptMerchantCoupons, elemAcceptGiftCertificates)
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createMerchantCalculations()"

    ' Verify that the elemMerchantCalculationsUrl parameter has a value
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "elemMerchantCalculationsUrl", _
        elemMerchantCalculationsUrl

    ' Define the variables used to create the <merchant-calculations> element
    Dim domMerchantCalculations
    Dim domMerchantCalculationsUrl
    Dim domAcceptMerchantCoupons
    Dim domAcceptGiftCertificates

    ' Create the <merchant-calculations> element
    Set domMerchantCalculationsObj = Server.CreateObject(strMsxmlDomDocument)
    domMerchantCalculationsObj.async = False
    Set domMerchantCalculations = domMerchantCalculationsObj.appendChild( _
        domMerchantCalculationsObj.createElement("merchant-calculations"))

    ' Create the <merchant-calculations-url> element
    Set domMerchantCalculationsUrl = domMerchantCalculations.appendChild( _
        domMerchantCalculationsObj.createElement("merchant-calculations-url"))
    domMerchantCalculationsUrl.Text = elemMerchantCalculationsUrl

    ' Create the <accepts-merchant-coupons> element
    If elemAcceptMerchantCoupons <> "" Then

        Set domAcceptMerchantCoupons = domMerchantCalculations.appendChild( _
            domMerchantCalculationsObj.createElement( _
                "accept-merchant-coupons"))

        domAcceptMerchantCoupons.Text = elemAcceptMerchantCoupons

    End If

    ' Create the <accepts-gift-certificates> element
    If elemAcceptGiftCertificates <> "" Then

        Set domAcceptGiftCertificates = domMerchantCalculations.appendChild( _
            domMerchantCalculationsObj.createElement( _
                "accept-gift-certificates"))

        domAcceptGiftCertificates.Text = elemAcceptGiftCertificates

    End If
    
    Set createMerchantCalculations = domMerchantCalculationsObj

    ' Release the variables used to create the <merchant-calculations> element
    Set domMerchantCalculations = Nothing
    Set domMerchantCalculationsUrl = Nothing
    Set domAcceptMerchantCoupons = Nothing
    Set domAcceptGiftCertificates = Nothing

End Function


'*******************************************************************************
' The createMerchantCheckoutFlowSupport function builds a
' <merchant-checkout-flow-support> XML DOM. This XML contains
' information about taxes, shipping and other custom calculations
' to be used in the checkout process. The XML also contains URLs
' used during the checkout process, such as URLs for the customer
' to edit a shipping cart or to continue shopping.
'
' Input:       elemEditCartUrl          URL to visit if the customer wants
'                                       to edit the shopping cart
' Input:       elemContinueShoppingUrl  URL to visit if the customer wants
'                                       to continue shopping
'
' Returns:     <merchant-checkout-flow-support> XML DOM
'*******************************************************************************
Function createMerchantCheckoutFlowSupport(elemEditCartUrl, _
    elemContinueShoppingUrl)

    ' Define objects used to create the <merchant-checkout-flow-support> XML
    Dim domMerchantCFS
    Dim domEditCartUrl
    Dim domContinueShoppingUrl
    Dim domShippingMethodsRoot
    Dim domTaxTablesRoot
    Dim domMerchantCalculationsRoot

    ' Create the <merchant-checkout-flow-support> element
    Set domMerchantCFSObj = Server.CreateObject(strMsxmlDomDocument)
    domMerchantCFSObj.async = False
    Set domMerchantCFS = domMerchantCFSObj.appendChild( _
        domMerchantCFSObj.createElement("merchant-checkout-flow-support"))

    ' Add the <edit-cart-url> element
    If elemEditCartUrl <> "" Then
        Set domEditCartUrl = domMerchantCFS.appendChild( _
            domMerchantCFSObj.createElement("edit-cart-url"))
        domEditCartUrl.Text = elemEditCartUrl
    End If
    
    ' Add the <continue-shopping-url> element
    If elemContinueShoppingUrl <> "" Then
        Set domContinueShoppingUrl = domMerchantCFS.appendChild( _
            domMerchantCFSObj.createElement("continue-shopping-url"))
        domContinueShoppingUrl.Text = elemContinueShoppingUrl
    End If

    ' Add the <shipping-methods> element
    If IsObject(domShippingMethodsObj) Then
        Set domShippingMethodsRoot = domShippingMethodsObj.documentElement
        domMerchantCFS.appendChild(domShippingMethodsRoot.cloneNode(true))
    End If

    ' Add the <tax-tables> element
    If IsObject(domTaxTablesObj) Then
        Set domTaxTablesRoot = domTaxTablesObj.documentElement
        domMerchantCFS.appendChild(domTaxTablesRoot.cloneNode(true))
    End If

    ' Add the <merchant-calculations> element
    If IsObject(domMerchantCalculationsObj) Then
        Set domMerchantCalculationsRoot = _
            domMerchantCalculationsObj.documentElement
        domMerchantCFS.appendChild(domMerchantCalculationsRoot.cloneNode(true))
    End If

    Set createMerchantCheckoutFlowSupport = domMerchantCFSObj

    ' Release objects used to create the <merchant-checkout-flow-support> XML
    Set domMerchantCFS = Nothing
    Set domEditCartUrl = Nothing
    Set domContinueShoppingUrl = Nothing
    Set domShippingMethodsObj = Nothing
    Set domShippingMethodsRoot = Nothing
    Set domTaxTablesObj = Nothing
    Set domTaxTablesRoot = Nothing
    Set domMerchantCalculationsObj = Nothing
    Set domMerchantCalculationsRoot = Nothing

End Function


'*******************************************************************************
' The createCheckoutShoppingCart function returns the
' <checkout-shopping-cart> XML DOM, which contains all of the items
' and checkout-related information for an order.
'
' Returns: <checkout-shopping-cart> XML
'*******************************************************************************
Function createCheckoutShoppingCart()
    
    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "createCheckoutShoppingCart()"

    ' Verify that there is a <shopping-cart> XML DOM and a
    ' <merchant-checkout-flow-support> XML DOM
    errorType = "MISSING_PARAM"
    If Not(IsObject(domShoppingCartObj)) Then
        errorHandler errorType, strFunctionName, "domShoppingCartObj", _
            domShoppingCartObj
    End If

    ' Define the variables used to create the <checkout-shopping-cart> element
    Dim domCheckoutShoppingCart
    Dim domShoppingCartRoot
    Dim domCFSRoot
    Dim domCFS
    Dim domMerchantCFSRoot
    
    ' Create the <checkout-flow-support> element and add
    ' the <merchant-checkout-flow-support> element as a child element
    Set domCFSObj = Server.CreateObject(strMsxmlDomDocument)
    domCFSObj.async = False
    Set domCFS = _
        domCFSObj.appendChild(domCFSObj.createElement("checkout-flow-support"))

    Set domMerchantCFSRoot = domMerchantCFSObj.documentElement
    domCFS.appendChild(domMerchantCFSRoot.cloneNode(true))

    Set domCheckoutShoppingCartObj = Server.CreateObject(strMsxmlDomDocument)
    domCheckoutShoppingCartObj.async = False

    domCheckoutShoppingCartObj.appendChild( _
        domCheckoutShoppingCartObj.createProcessingInstruction( _
            "xml", strXmlVersionEncoding))

    ' Create the <checkout-shopping-cart> element
    Set domCheckoutShoppingCart = domCheckoutShoppingCartObj.appendChild( _
        domCheckoutShoppingCartObj.createElement("checkout-shopping-cart"))
    domCheckoutShoppingCart.setAttribute "xmlns", strXmlns

    ' Add the <shopping-cart> element as a child element of the
    ' <checkout-shopping-cart> element
    Set domShoppingCartRoot = domShoppingCartObj.documentElement
    domCheckoutShoppingCart.appendChild(domShoppingCartRoot.cloneNode(true))

    Set domCFSRoot = domCFSObj.documentElement
    domCheckoutShoppingCart.appendChild(domCFSRoot.cloneNode(true))

    createCheckoutShoppingCart = domCheckoutShoppingCartObj.xml

    ' Release the variables used to create the <checkout-shopping-cart> element
    Set domShoppingCartObj = Nothing
    Set domCheckoutShoppingCart = Nothing
    Set domShoppingCartRoot = Nothing
    Set domCFSObj = Nothing
    Set domCFSRoot = Nothing
    Set domCFS = Nothing
    Set domMerchantCFSObj = Nothing
    Set domMerchantCFSRoot = Nothing

End Function

%>