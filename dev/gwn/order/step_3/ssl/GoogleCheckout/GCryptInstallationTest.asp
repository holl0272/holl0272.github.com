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


'*************************************************
' BEFORE USING PLEASE READ
'
' Before using this installation test script, you
' must install the ActiveX Com object onto your
' Windows Server.
' The COM library has been packaged as a .MSI file
' and will automatically install into the 
' Windows Component Services COM+ packages
' You will see it show up as "GCrypt"
'
' After running the .MSI installation, copy this
' installation test .asp page to your webserver
' and simply call the page
'*************************************************

Dim str
Dim key
Dim b64str
Dim b64signature

str = "test_data"
key = "test_key"


    Set cryptobj = Server.CreateObject("GCrypt.g_crypt.2")
    
    ' BASE64-ENCODE (input data as String) 
	' - outputs base64 encoded binary string
    b64str = cryptObj.base64Encode(str)

    ' HMACSHA1(input data as String, secret key as String)
	' - outputs base64 encoded binary string
    b64signature = cryptobj.generateSignature(str,key)
    
%>
<html>
<head>
    <title>GCrypt Installation Test Page</title>
</head>

<body>
    <center>
        <h1>HMAC-SHA1 and Base64-Encode Test Page</h1>
    </center>

    <table border='1' width="800">
        <tr>
            <td>String used for calculations:</td>

            <td>test_data</td>
        </tr>
        <tr>
            <td>String used as key:</td>
            <td>test_key</td>
        </tr>

        <tr>
            <td colspan="2">&nbsp;</td>
        </tr>

        <tr>
            <td>Base64 Encoded String Calculated by GCrypt</td>

            <td><%=b64str%></td>
        </tr>

        <tr>
            <td>Expected Base64 Encoding of the String</td>

            <td>dGVzdF9kYXRh</td>
        </tr>

        <tr>
            <td colspan='2'>&nbsp;</td>
        </tr>

        <tr>
            <td>Base64 Encoding of HMAC-SHA1 Calculated by GCrypt</td>

            <td><%=b64signature%></td>
        </tr>

        <tr>
            <td>Expected Base64 Encoding of HMAC-SHA1 Signature</td>

            <td>2G2Jv74LpjS1Emu6tJRrTES/Dlw=</td>
        </tr>

    </table>
</body>
</html>
