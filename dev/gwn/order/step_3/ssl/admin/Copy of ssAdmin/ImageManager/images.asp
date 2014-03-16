<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
<!--#include file="SOSLibrary/incPureUpload.asp" -->
<%
If Request.Form("butDeleteImage") <> "" Then

	'Create the FileSystemObject
	Dim objDelete, delPos
	Set objDelete = Server.CreateObject("Scripting.FileSystemObject")
	delPos = Instr(Request.Form("selectImg"),"/")
	'Delete the file
	objDelete.DeleteFile Server.MapPath(strCurrentFolder & Right(Request.Form("selectImg"),Len(Request.Form("selectImg"))-delPos)), False
	
	'Redirect to clear all form variables.
	Response.Redirect("images.asp?" & Request.QueryString)
End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
    "http://www.w3.org/TR/html4/loose.dtd">
<html>
<!-- Page Design by Web Shorts Site Design, www.web-shorts.com -->
<head>
<title>Choose a Product Image</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

  function pickImage() {
    if (self.opener && MM_findObj('selectImg')){
      self.opener.document.frmData.<%= Request.Querystring("img") %>.value='<%= strPathFromProducts %>'+MM_findObj('selectImg').options[MM_findObj('selectImg').selectedIndex].value;
      }
    self.close();
  return true;
 }

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function GP_popupConfirmMsg(msg) { //v1.0
  document.MM_returnValue = confirm(msg);
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
<link href="imgmgr.css" rel="stylesheet" type="text/css">
</head>
<body<% If Request.Querystring("path") <> "" Then Response.Write(" onLoad=""MM_swapImage('ProductImage','','" & strAbsoluteURL & Right(Request.Querystring("path"),Len(Request.QueryString("path"))-Len(strPathFromProducts)) & "',1)""") %>>
<% If vDebug = 1 Then DebugValues %><form name="form1" method="post" action="images.asp?<%= Request.Querystring %>">
  <table border="0" align="center" cellpadding="1" cellspacing="0" class="border">
    <tr> 
      <td> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr align="center"> 
            <th colspan="2"><b>Pick Image</b></th>
          </tr>
          <tr> 
            <td colspan="2" valign="top" class="matte">Choose a Product <%= Request.Querystring("img") %> Image. Click Select Image to add the selected image 
              path to the product page. Click Upload Image to add a new image 
              to the list. Click Delete Image to delete the selected image.</td>
          </tr>
          <tr> 
            <td colspan="2" valign="top" class="matte"> 
<table border="0" cellpadding="2" cellspacing="0">
<tr valign="top"> 
                  <td> 
                    <% If ShowSubs = True then %>
                    Folders:<br>
                    <select name="showFolder" onChange="MM_goToURL('parent','images.asp?img=<%= Request.QueryString("img") %>&path=<%= Request.Querystring("path") %>&show='+MM_findObj('showFolder').options[MM_findObj('showFolder').selectedIndex].value,1);return document.MM_returnValue">
<%
	Dim objFolderFSO
	Set objFolderFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	Dim objFolder
	Set objFolder = objFolderFSO.GetFolder(Server.MapPath(strCurrentFolder))
	Dim objFolderRoot, oFolderPos
	objFolderRoot = objFolder
	oFolderPos = Len(objFolderRoot)
	ListFolders objFolder

Sub ListFolders(objFolder)

Dim trimLength
Dim pstrOut

	pstrOut = "<option value=""" & Server.URLEncode(objFolder) & """"
	If CStr(Request.QueryString("show")) = CStr(objFolder) Then pstrOut = pstrOut & " SELECTED"
	pstrOut = pstrOut & ">"

	If objFolder = objFolderRoot Then 
		pstrOut = pstrOut & "Root Image Folder"
	Else
		trimLength = Len(objFolder) - oFolderPos - 1
		pstrOut = pstrOut & "/" & Replace(Right(objfolder,trimLength),"\","/")
	End If
	pstrOut = pstrOut & "</option>" & VbCrLf
	
	If inStr(1, objfolder, "_vti_") < 1 Then Response.Write pstrOut

	'Now, use a for each...next to loop through the Files collection
	If ShowSubs = True Then
		Dim objSubFolder
		For Each objSubFolder in objFolder.SubFolders
			ListFolders objSubFolder
		Next
	End If
	
End Sub

%>
                    </select>
                    <br><% End If %>
                      <%
'=================================
'Directory and page options
'=================================
Const DispPageName = True
Const PageExtensions = "gif,jpg"

'=================================
'Work Code - Be careful editing below this line.
'=================================
Dim curFolder
	curFolder = Left(strCurrentFolder,Len(strCurrentFolder)-1)
	curFolder = Right(curFolder,Len(curFolder) - inStrRev(curFolder,"/"))

Dim objFSO, MyFolder
If Request.QueryString("show") <> "" Then
  	MyFolder = Request.QueryString("show")
Else
	MyFolder = Server.MapPath(strCurrentFolder)
End If
  
'Is the directory valid?
If Right(MyFolder,1) <> "\" Then MyFolder = Myfolder & "\"
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  
'If directory does not exist, return an error.
If objFSO.FolderExists(MyFolder) = False Then
	Response.Write "Invalid folder: " & MyFolder
	Response.End
End If

'Obtain an folder object instance for a particular directory
Set objFSO = objFSO.GetFolder(MyFolder)
  
'Use a For Each ... Next loop to display the files
Dim objFile, fso, f, ts, myText, fts, strExtension, allowedExtensions, i
Dim FolderRoot, oPos, imgCount
FolderRoot = objFSO
oPos = Len(Server.MapPath(strCurrentFolder))
imgCount = ImageCount(objFSO)
If imgCount <> 0 Then %>
<select name="selectImg" size="10" ondblclick="MM_callJS('pickImage()')" onChange="MM_swapImage('ProductImage','','<%= strAbsoluteURL %>'+MM_findObj('selectImg').options[MM_findObj('selectImg').selectedIndex].value,1)">
	<% ListImages(objFSO) %>
</select>
<%
Else
	Response.Write "<p>No images in the current folder.</p>"
End If
  
Sub ListImages(objFSO)
	Dim TrimLength, myPath, imgCount
	imgCount = 0
	For Each objFile in objFSO.Files
	
		set fso = server.CreateObject("Scripting.FileSystemObject") 
		set f = fso.GetFile(objFile.Path) 
		strExtension = Ucase(Right(f, Len(f) - InStrRev(f, ".")))
		myText = ""
		allowedExtensions = Split(PageExtensions,",")
		For i = 0 to UBound(allowedExtensions)
			If strExtension = UCase(allowedExtensions(i)) Then
				trimLength = Len(objFSO) - oPos
				myPath = Right(objFSO,trimLength)
				myPath = Replace(myPath,"\","/")
				Response.Write ("<option value=""" & curFolder & myPath & "/" & objFile.Name & """")
				If Request.Querystring("path") = curFolder & myPath & "/" & objFile.Name Then Response.Write " selected"
				Response.Write(">")
				If myPath = "" Then
					Response.Write(objFile.Name)
				Else
					Response.Write(objFile.Name)
				End If
				Response.Write "</option>" & VbCrLf
				imgCount = imgCount+1
			End If
		Next
	Next
	If imgCount = 0 Then Response.Write "<option>No Images</option>"
	'If ShowSubs = True then
	'	Dim SubFolder
	'	For Each SubFolder in objFSO.SubFolders
	'		ListImages SubFolder
	'	Next
	'End If
End Sub

Function ImageCount(objFSO)
	Dim TrimLength, myPath, imgCount
	imgCount = 0
	For Each objFile in objFSO.Files
	
		set fso = server.CreateObject("Scripting.FileSystemObject") 
		set f = fso.GetFile(objFile.Path) 
		strExtension = Ucase(Right(f, Len(f) - InStrRev(f, ".")))
		myText = ""
		allowedExtensions = Split(PageExtensions,",")
		For i = 0 to UBound(allowedExtensions)
			If strExtension = UCase(allowedExtensions(i)) Then
				imgCount = imgCount+1
			End If
			If imgCount > 0 Then 
				ImageCount = imgCount
				Exit Function
			End If
		Next
	Next
	ImageCount = imgCount
End Function

%>
                    <br>
                    <% If imgCount <> 0 Then %>
                    <input type="button" name="pickImage" value="Select Image" onclick="MM_callJS('pickImage()')" class="go" style="width: 150px;">
                    <br><% End If %>
                    <input type="submit" name="butUploadImage" value="Upload/Manage" class="go" onclick="MM_goToURL('parent','imageupload.asp?<%= Request.QueryString %>');return document.MM_returnValue" style="width: 150px;">
                    <% If ImgCount <> 0 Then %><br>
                    <input type="submit" name="butDeleteImage" value="Delete Image" class="go" onclick="GP_popupConfirmMsg('Are you sure you want to delete the selected image? This action cannot be undone. Be sure no products are currently using this image.');return document.MM_returnValue" style="width: 150px;">
<% End If %>
                  </td>
                  <td> 
                    <p><img src="transparent.gif" alt="Product Image" name="ProductImage" border="0" ondblclick="MM_callJS('pickImage()')"></p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <th colspan="2" align="center"><b><a href="javascript:window.close();">Close</a></b></th>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
