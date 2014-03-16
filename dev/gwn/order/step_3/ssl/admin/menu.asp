<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<title>StoreFront - Web Store Adminstrative Menu</title>
<script language="JavaScript">
<!-- Hide the script from old browsers --
function loadalert ()
         {alert("Your Administrative web is unsecured. Please secure the web before publishing in order to protect your store and sales information")}
// --End Hiding Here -->
</script>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<!--webbot bot="HTMLMarkup" StartSpan -->
<% If Request.ServerVariables("LOGON_USER") = "" Then %>
<body onLoad="loadalert()">
<% End If %>
<!--webbot BOT="HTMLMarkup" endspan -->



<p align="center"><img src="../../images/sflogo.gif" alt="sflogo.gif (13030 bytes)" WIDTH="300" HEIGHT="171"><br>
<u><big><big><strong>Store Administration Services</strong></big></big></u></p>
<div align="center"><center>

  <table cellpadding="0" cellspacing="0" height="300">
    <tr>
      <td align="right" colspan="2" height="24" bgcolor="#A09A8B"><p align="center"><big><strong>Product
    Management</strong></big></p>
      </td>
    </tr>
    <tr>
      <td width="50%" height="21" colspan="2" align="center"><a href="prodadd.htm">Add New
    Product Items </a></td>
    </tr>
    <tr>
      <td width="50%" height="21" align="center" colspan="2"><a href="proddelete.htm">Delete
    Products </a></td>
    </tr>
    <tr>
      <td width="50%" height="21" align="center" colspan="2"><a href="prodedit.htm">Edit
    Products</a></td>
    </tr>
    <tr>
      <td width="50%" height="21" align="center" colspan="2"><a href="prodlist.asp">List All
    Products</a></td>
    </tr>
    <tr>
      <td align="right" height="21"></td>
      <td width="50%" height="21"></td>
    </tr>
    <tr>
      <td align="right" colspan="2" height="24" bgcolor="#A09A8B"><p align="center"><big><strong>Sales
    Reports</strong></big></p>
      </td>
    </tr>
    <tr>
      <td align="right" colspan="2" height="24"><big><strong><p align="center"></strong></big><a href="reports.htm#Detail">Transaction Detail</a></p>
      </td>
    </tr>
    <tr>
      <td align="right" colspan="2" height="24"><big><strong><p align="center"></strong></big><a href="reports.htm#Summary">Sales Summary Report</a></p>
      </td>
    </tr>
    <tr>
      <td align="right" colspan="2" height="24"><big><strong><p align="center"></strong></big><a href="reports.htm#CyberCash">Transaction Service Reports</a></p>
      </td>
    </tr>
    <tr>
      <td align="right" height="21"></td>
      <td width="50%" height="21"></td>
    </tr>
    <tr>
      <td align="right" height="21" colspan="2" bgcolor="#A09A8B"><p align="center"><big><strong>Store
    Management</strong></big></p>
      </td>
    </tr>
    <tr>
      <td align="right" height="21" colspan="2"><p align="center"><a href="set_up.asp?Update=0">New
    Store Set-Up</a></p>
      </td>
    </tr>
    <tr>
      <td align="right" height="21"></td>
      <td width="50%" height="21"></td>
    </tr>
  </table>
  </center></div>

<p>&nbsp; </p>
</body>
</html>
