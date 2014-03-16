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
' "GlobalAPIFunctions.asp" contains a set of functions and variables that 
' are frequently used throughout the sample code.
' 
' You should also look at the Demo files to learn more about how to call
' each function and what it returns.
'*******************************************************************************

OPTION Explicit

Dim strMerchantId
Dim strMerchantKey
Dim strMsxmlDomDocument
Dim strXmlns
Dim strXmlVersionEncoding
Dim cryptObj
Dim attrCurrency
Dim logFilename
Dim baseUrl
Dim checkoutUrl
Dim checkoutDiagnoseUrl
Dim requestUrl
Dim requestDiagnoseUrl
Dim errorReportType

setGlobalVariables

' Set global variables and constants
Function setGlobalVariables

    ' +++ CHANGE ME +++
    ' The logFilename variable identifies the file where messages
    ' will be logged. You can change the variable's value to change
    ' the log file's location.
    logFilename = "log.out"

    ' 
    ' +++ CHANGE ME +++
    ' The errorReportType variable specifies the
    ' manner in which errors will be reported. There are three
    ' possible values:
    ' 1 = Log the error message to the IIS log file
    ' 2 = Display the error message in the browser
    ' 3 = Log the error message to the IIS log file
    '         and also display it in the browser
    ' 
    ' Error messages are for debugging purposes only. When you are
    ' done with integration, change the errorReportType variable to 1
    ' so that no error messages will be displayed to the end user.
    ' 
    errorReportType = "3"

    ' +++ CHANGE ME +++
    ' The attrcurrency variable specifies a default currency that
    ' is used in several places throughout the ASP libraries. You
    ' will need to update this value if you sell products in
    ' currencies other than U.S. dollars. The variable's value should
    ' be a three-letter ISO 4217 currency code:
    ' http://www.iso.org/en/prods-services/popstds/currencycodeslist.html
    '
    ' If you sell products in multiple currencies, you may need to
    ' implement a function that returns the appropriate currency code
    ' for each user.
    '
    ' Note: Google Checkout only supports USD at this time.
    attrCurrency = "USD"

    ' This constant identifies the location of the Google Checkout XML schema
    strXmlns = "http://checkout.google.com/schema/2"

    ' These two function calls set global variables for your
    ' merchant ID and merchant key
    strMerchantId = getMerchantId
    strMerchantKey = getMerchantKey

    ' These constants specify the URLs to which Google Checkout API requests 
    ' are sent
    ' +++ CHANGE ME +++
    ' Please remember that your production systems must send requests to 
    ' https://checkout.google.com
    baseUrl = "https://sandbox.google.com/cws/v2/Merchant/" & strMerchantId
    checkoutUrl = baseUrl & "/checkout"
    checkoutDiagnoseUrl = baseUrl & "/checkout/diagnose"
    requestUrl = baseUrl & "/request"
    requestDiagnoseUrl = baseUrl & "/request/diagnose"

    ' DomDocument Version
    strMsxmlDomDocument = "Msxml2.DOMDocument.3.0"

    ' XML Version and Encoding info
    strXmlVersionEncoding = "version=""1.0"" encoding=""UTF-8"""

    ' Cryptography COM Object
    Set cryptObj = Server.CreateObject("GCrypt.g_crypt.2")

End Function

'*******************************************************************************
' The getMerchantId function securely fetches and returns your Merchant ID.
' 
' Returns:  Merchant Id
'*******************************************************************************
Function getMerchantId()

    ' +++ CHANGE ME +++
    ' Please set the return value to your Google Checkout merchant ID.
    ' This change is mandatory or this code will not work.

    getMerchantId = ""

End Function


'*******************************************************************************
' The getMerchantKey function securely fetches and returns your Merchant Key.
' 
' Returns:  Merchant Key
'*******************************************************************************
Function getMerchantKey()

    ' +++ CHANGE ME +++
    ' Please set the return value to your Google Checkout merchant key.
    ' This change is mandatory or this code will not work.
    getMerchantKey = ""

End Function


'*******************************************************************************
' The SendRequest function verifies that you have provided values for
' all of the parameters needed to send a Google Checkout
' Checkout or Order Processing API request. It then logs the request, 
' sets HTTP headers and executes the request, and logs the response.
'
' Input:      request       XML API request
'             strPostUrl    URL address to which the request will be sent
'             response      synchronous response from the Google Checkout
'                               server
' Returns:    XML response from the Google server as text
'*******************************************************************************
Function sendRequest(request, strPostUrl)

    ' Check for errors
    Dim strFunctionName
    Dim errorType

    strFunctionName = "sendRequest()"

    ' Check for missing parameters
    errorType = "MISSING_PARAM"
    checkForError errorType, strFunctionName, "request", request
    checkForError errorType, strFunctionName, "strPostUrl", strPostUrl
    checkForError errorType, strFunctionName, "strMerchantId", strMerchantId
    checkForError errorType, strFunctionName, "strMerchantKey", strMerchantKey

    ' Define objects used to send the HTTP request
    Dim xmlHttp
    Dim strAuthentication 
    Dim bCheckout

    ' Log the outgoing message
    logMessage logFilename, request

    ' Create the XMLHttpRequest object
    Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")

    ' The HTTP request method is POST
    xmlHttp.open "POST", strPostUrl, False

    ' Do NOT ignore Server SSL Cert Errors
    Const SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2
    Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
    xmlHttp.setOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, _
        (xmlHttp.getOption(SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS) - _
        SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)

    bCheckout = InStr(strPostUrl, "checkout")

    ' This block executes if this is a Checkout API request
    If bCheckout <> 0 Then

        ' Set HTTP header
        xmlHttp.setRequestHeader "Content-Type", _
            "application/x-www-form-urlencoded"

    ' This block executes if this is not a Checkout API request
    Else

        ' Build HTTP Basic Authentication scheme
        strAuthentication = createHttpBasicAuthentication(strMerchantId, _
            strMerchantKey)

        ' Set HTTP headers
        xmlHttp.SetRequestHeader "Authorization", strAuthentication
        xmlHttp.SetRequestHeader "Content-Type", "application/xml"
        xmlHttp.SetRequestHeader "Accept", "application/xml"
    
    End If

    ' Transmit the request
    xmlHttp.send request

    ' Log the HTTP response
    logMessage logFilename, xmlHttp.responseText

    ' Return the response from the Google server
    sendRequest = xmlHttp.responseText

    ' Release the object used to send the request
    Set xmlHttp = Nothing

End Function


'*******************************************************************************
' The createHttpBasicAuthentication creates a string in the format
'      merchantId:merchantKey
' and then base64 encodes that string. This string is used to send
' Google Checkout API requests that are not Checkout API requests.
'
' Input:      strMerchantId            Your Merchant ID
'             strMerchantKey           Your secret Merchant Key
' Returns:    HTTP Basic authentication string
'*******************************************************************************
Function createHttpBasicAuthentication(strMerchantId, strMerchantKey)

    Dim strCredential
    Dim b64credential
    Dim strAuthentication

    ' Create "userid:password" 
    strCredential = strMerchantId & ":" & strMerchantKey

    ' Base64-encode "userid:password"
    b64credential = cryptObj.base64Encode(strCredential)

    ' Create "Basic dXNlcmlkOnBhc3N3b3Jk"
    strAuthentication = "Basic " & b64credential

    ' Return the HTTP Basic Authentication string
    createHttpBasicAuthentication = strAuthentication

End Function


'******************************************************************************
' The displayDiagnoseResponse function is a debugging function that
' sends a Google Checkout API request and then evaluates the Google Checkout 
' response to determine whether the request used valid XML. If the request did
' not use valid XML, the function displays an error message and a link
' where you can edit the XML and then try to validate it again.
'
' This function calls the SendRequest function to execute the API request.
'
' Input:       request         XML API request
' Input:       strPostUrl      URL address to which the form should be posted
' Input:       xml             Unencoded version of XML used in API request
' Input:       action          This variable indicates whether the function 
'                              should print a form on the page containing 
'                              information about the API request if the XML 
'                              is invalid.
' Returns:     response        Boolean (true=XML is valid;false=XML is invalid)
'******************************************************************************

Function displayDiagnoseResponse(request, strPostUrl, xml, action)

    ' Define objects used to diagnose the API response
    Dim diagnoseResponse
    Dim bValidated
    Dim domResponse
    Dim strRootTag
    Dim nodeList
    Dim strResult

    ' Execute the API request and capture the Google Checkout server's response
    diagnoseResponse = sendRequest(request, strPostUrl)

    ' If the function finds that the request contained valid XML, the
    ' $validated variable will be set to true
    bValidated = false

    Set domResponse = Server.CreateObject("Msxml2.DOMDocument.3.0")
    domResponse.loadXml diagnoseResponse

    strRootTag = domResponse.documentElement.tagName

    ' This if-else block determines whether the API response indicates
    ' that the response contained invalid XML or if there was some other
    ' problem associated with the request, such as an invalid signature.
    If strRootTag = "diagnosis" Then
        Set nodeList = _
            domResponse.documentElement.getElementsByTagName("string")
        If nodeList.length > 0 Then
            strResult = nodeList(0).text
        Else
            bValidated = True
        End If
    Elseif strRootTag = "error" Then
        Set nodeList = _
            domResponse.documentElement.getElementsByTagName("error-message")
        strResult = nodeList(0).text
    ElseIf strRootTag = "request-received" Then
        bValidated = true
    End If

    ' If the request is invalid, print the reason that the request is
    ' invalid if the errorReportType variable indicates that errors
    ' should be displayed in the user's browser. Also display a link 
    ' to a tool where the user can edit the XML request unless the
    ' validation request was submitted from that tool.
    If bValidated = False And (errorReportType = 2 Or errorReportType = 3) Then
        Response.write "<tr><td style=""color:red""><p>" & _
            "<span style=""text-align:center""><h2>" & _
            "This XML is NOT Validated!</h2></span></p>"
        Response.write "<p style=""text-align:left""><b>" & _
            Server.HTMLEncode(strResult) & "</b></p>"
        If action = "debug" Then
            Response.write "<p><form method=POST action=DebuggingTool.asp>"
            Response.write "<input type=""hidden"" name=""xml"" value=""" & _
                Server.HTMLEncode(xml) & """/>"
            Response.write "<input type=""hidden"" name=""toolType"" " & _
                "value=""Validate XML""/>"
            Response.write "<input type=""submit"" name=""Debug"" " & _
                "value=""Debug XML""/>"
            Response.write "</form></p></td></tr>"
        End If
    End If

    ' Return a Boolean value indicating whether the request
    ' contained valid XML.
    displayDiagnoseResponse = bValidated
End Function


'*******************************************************************************
' The CheckForError function determines whether a parameter has a null
' value and prints the appropriate error message if the parameter does
' have a null value.
'
' Input:       errorType          The type of error being flagged.
'                                     e.g. MISSING_PARAM
' Input:       strFunctionName    The function where the error occurred
' Input:       strParamName       The name of the parameter being checked
' Input:       strParamValue      The parameter value submitted to the function
'*******************************************************************************
Function checkForError(errorType, strFunctionName, strParamName, strParamValue)
    If strParamValue = "" Then
        errorHandler errorType, strFunctionName, strParamName, strParamValue
    End If
End Function


'*******************************************************************************
' The errorHandler function returns the error message that should be
' logged for the $error_type.
'
' Input:       errorType            The type of error being flagged.
'                                     e.g. "MISSING_PARAM",
'                                     "INVALID_INPUT_ARRAY", "MISSING_CURRENCY"
'                                     "MISSING_TRACKING"
' Input:       errorFunctionName    The function where the error occurred
' Input:       errorParamName       The name of the parameter being checked
' Input:       errorParamValue      The parameter value submitted to the function
' Returns:     error message
'*******************************************************************************
Function errorHandler(errorType, errorFunctionName, errorParamName, _
    errorParamValue) 

    Dim errstr

    Select Case errorType 

        ' MISSING_PARAM error
        ' A function call omits a required parameter.
        Case "MISSING_PARAM"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: """ & errorParamName _
                & """ must be provided."

        ' MISSING_PARAM_NONE error
        ' A function call must have a value for at least one parameter.
        Case "MISSING_PARAM_NONE"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: " _
                & "At least one parameter should be provided."

        ' INVALID_INPUT_ARRAY error
        ' AddAreas() function called with invalid value for
        ' $state_areas or $zip_areas parameter
        Case "INVALID_INPUT_ARRAY"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Invalid Input: """ & errorParamName _
                & """ should be an array."

        ' MISSING_CURRENCY error
        ' The attrCurrency value is empty.
        Case "MISSING_CURRENCY"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: ""attrCurrency"" " _
                & "should be set when the ""elemAmount"" is set."

        ' MISSING_TRACKING error
        ' The ChangeShippingInfo() function in
        ' OrderProcessingAPIFunctions.asp is being called without
        ' specifying a tracking number even though a shipping
        ' carrier is specified.
        Case "MISSING_TRACKING"
            errstr = "Error calling Function """ & errorFunctionName _
                & """: Missing Parameter: ""elemTrackingNumber"" " _
                & "should be set when the ""elemCarrier"" is set."

        Case Else

    End Select 

    ' Print the error message to the screen
    If (errorReportType = 2) Or (errorReportType = 3) Then 

        Dim errstrHtml
        errstrHtml = errstr & "<br><br>"

        Response.write errstrHtml

    End If

    ' Write out the error message to the IIS Log File
    If (errorReportType = 1) Or (errorReportType = 3) Then 

        Response.appendToLog errstr

    End If

    Response.End

End Function



'*******************************************************************************
'  The logMessage function logs a message to a local file. The function
'  also logs the time that the message is logged.
'  
'  Input:  logFilename      The filename to which the message should be logged
'  Input:  message          The message to be logged
'*******************************************************************************
Function logMessage(logFilename, message)

    Dim oFs
    Dim oTextFile

    ' Print out the notification message to a local file
    Set oFs = Server.createobject("Scripting.FileSystemObject")
    Const ioMode = 8
    Set oTextFile = oFs.openTextFile(logFilename, ioMode, True)
    oTextFile.writeLine now
    oTextFile.writeLine message
    oTextFile.close

    ' Free object
    Set oTextFile = Nothing
    Set oFS = Nothing

End Function


%>