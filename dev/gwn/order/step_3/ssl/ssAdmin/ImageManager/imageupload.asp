<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="SOSLibrary/incPureUpload.asp" -->
<%
'*** Pure ASP File Upload -----------------------------------------------------
' Copyright 2000 (c) George Petrov, www.UDzone.com
' Process the upload
' Version: 2.0.3
'------------------------------------------------------------------------------
'*** File Upload to: "../../images", Extensions: "GIF,JPG", Form: form1, Redirect: "images.asp", "file", "", "error", "true", "", "" , "", "", "", "", "600", "fileCopyProgress.htm", "300", "100"

Dim UploadQueryString, GP_uploadAction, UploadRequest
PureUploadSetup

If (CStr(Request.QueryString("GP_upload")) <> "") Then
  Server.ScriptTimeout = 600
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Set UploadRequest = CreateObject("Scripting.Dictionary") 
  BuildUploadRequest RequestBin, "emptyfolder", "file", MaxPhysicalSize, "error"
  GP_redirectPage = "imageupload.asp?img=" & Request.QueryString("img") & "&show=" & Server.URLEncode(UploadFormRequest("myField")) & "&fldraction=Image%20" & UploadFormRequest("file") & "%20Uploaded"
  
  If (GP_redirectPage <> "") Then
    If (InStr(1, GP_redirectPage, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      GP_redirectPage = GP_redirectPage & "?" & UploadQueryString
    End If
    Response.Redirect(GP_redirectPage)  
  end if  
end if  
if UploadQueryString <> "" then
  UploadQueryString = UploadQueryString & "&GP_upload=true"
else  
  UploadQueryString = "GP_upload=true"
end if  

' End Pure Upload
'------------------------------------------------------------------------------

If CStr(Request.Querystring("fldrmgmt")) = "true" Then
	Dim actionStatus, actionRedirect
	If Request.Form("sbtAddFolder") <> "" Then
		'Create an instance of the FileSystemObject
		Dim objCreateFSO
		Set objCreateFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		'Create C:\FooBar
		If Not objCreateFSO.FolderExists(Request.Form("lstParentFolder") & "\" & Request.Form("txtFldrName")) then
		  objCreateFSO.CreateFolder(Request.Form("lstParentFolder") & "\" & Request.Form("txtFldrName"))
		End If
		actionStatus = Server.URLEncode(Request.Form("txtFldrName") & " successfully added.")
		actionRedirect = ("imageupload.asp?img=" & Request.QueryString("img") & "&path=" & Request.QueryString("path") & "&show=" & Server.URLEncode(Request.QueryString("show")))

	End If
	If Request.Form("sbtDeleteFolders") <> "" Then
		'Create an instance of the FileSystemObject
		Dim objDeleteFSO
		Set objDeleteFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		'Delete C:\FooBar
		If objDeleteFSO.FolderExists(Request.Form("lstFldrName")) then
		  objDeleteFSO.DeleteFolder(Request.Form("lstFldrName"))
		End If
		actionStatus = Server.URLEncode(Request.Form("lstFldrName") & " and all images deleted.")
		actionRedirect = ("imageupload.asp?img=" & Request.QueryString("img") & "&path=" & Request.QueryString("path"))

	End If
	Response.Redirect(actionRedirect & "&fldraction=" & actionStatus)

End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
    "http://www.w3.org/TR/html4/loose.dtd">
<html>
<!-- Page Design by Web Shorts Site Design, www.web-shorts.com 
StoreFront Merchant Tools Image Management Mod 1.0-->
<head>
<title>Upload Image</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function checkFileUpload(form,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.0
  document.MM_returnValue = true;
  if (extensions != '') var re = new RegExp("\.(" + extensions.replace(/,/gi,"|") + ")$","i");
  for (var i = 0; i<form.elements.length; i++) {
    field = form.elements[i];
    if (field.type.toUpperCase() != 'FILE') continue;
    if (field.value == '') {
      if (requireUpload) {alert('File is required!');document.MM_returnValue = false;field.focus();break;}
    } else {
      if(extensions != '' && !re.test(field.value)) {
        alert('This file type is not allowed for uploading.\nOnly the following file extensions are allowed: ' + extensions + '.\nPlease select another file and try again.');
        document.MM_returnValue = false;field.focus();break;
      }
    document.PU_uploadForm = form;
    re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
    if(re.test(field.value) && (sizeLimit != '' || minWidth != '' || minHeight != '' || maxWidth != '' || maxHeight != '' || saveWidth != '' || saveHeight != '')) {
      checkImageDimensions(field,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
    } } }
}

function showImageDimensions(fieldImg) { //v2.0
  var isNS6 = (!document.all && document.getElementById ? true : false);
  var img = (fieldImg && !isNS6 ? fieldImg : this);
  if ((img.minWidth != '' && img.minWidth > img.width) || (img.minHeight != '' && img.minHeight > img.height)) {
    alert('Uploaded Image is too small!\nShould be at least ' + img.minWidth + ' x ' + img.minHeight); return;}
  if ((img.maxWidth != '' && img.width > img.maxWidth) || (img.maxHeight != '' && img.height > img.maxHeight)) {
    alert('Uploaded Image is too big!\nShould be max ' + img.maxWidth + ' x ' + img.maxHeight); return;}
  if (img.sizeLimit != '' && img.fileSize > img.sizeLimit) {
    alert('Uploaded Image File Size is too big!\nShould be max ' + (img.sizeLimit/1024) + ' KBytes'); return;}
  if (img.saveWidth != '') document.PU_uploadForm[img.saveWidth].value = img.width;
  if (img.saveHeight != '') document.PU_uploadForm[img.saveHeight].value = img.height;
  document.MM_returnValue = true;
}

function checkImageDimensions(field,sizeL,minW,minH,maxW,maxH,saveW,saveH) { //v2.0
  if (!document.layers) {
    var isNS6 = (!document.all && document.getElementById ? true : false);
    document.MM_returnValue = false; var imgURL = 'file:///' + field.value.replace(/\\/gi,'/');
    if (!field.gp_img || (field.gp_img && field.gp_img.src != imgURL) || isNS6) {field.gp_img = new Image();
		   with (field) {gp_img.sizeLimit = sizeL*1024; gp_img.minWidth = minW; gp_img.minHeight = minH; gp_img.maxWidth = maxW; gp_img.maxHeight = maxH;
  	   gp_img.saveWidth = saveW; gp_img.saveHeight = saveH; gp_img.onload = showImageDimensions; gp_img.src = imgURL; }
	 } else showImageDimensions(field.gp_img);}
}

function showProgressWindow(progressFile,popWidth,popHeight) { //v2.0
  if (document.MM_returnValue) {
    var w = 480, h = 340;
    if (document.all || document.layers || document.getElementById) {
      w = screen.availWidth; h = screen.availHeight;}
    var leftPos = (w-popWidth)/2, topPos = (h-popHeight)/2;
    document.progressWindow = window.open(progressFile,'ProgressWindow','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,width=' + popWidth + ',height='+popHeight);
    document.progressWindow.moveTo(leftPos, topPos);document.progressWindow.focus();
		window.onunload = function () {document.progressWindow.close();};
} }

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function GP_popupConfirmMsg(msg) { //v1.0
  document.MM_returnValue = confirm(msg);
}
//-->
</script>
<link href="imgmgr.css" rel="stylesheet" type="text/css">
</head>
<body>
<% If vDebug = 1 Then DebugValues %><table border="0" align="center" cellpadding="1" cellspacing="0" class="border">
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
<tr align="center"> 
          <th>Upload Image</th>
        </tr>
        <tr> 
          <td colspan="2" valign="top" class="matte">Browse to a file on your 
            local machine to upload. Be sure that the file is already optimized 
            for the web, and will fit inside your site layout. After the file 
            has finished uploading, you'll be returned to the image chooser. Use 
            the Folder Management section to add or remove folders. Be careful, 
            as removing a folder will delete all images inside that folder. </td>
        </tr>
        <tr> 
          <td colspan="2" valign="top" class="matte"> <br />
            <% If Request.QueryString("fldraction") <> "" Then %>
            <p align="center"><strong><%= Request.QueryString("fldraction")%></strong></p>
			<% End If %><table border="0" align="center" cellpadding="1" cellspacing="0" class="border">
<tr> 
                <td>
                  <table width="100%" border="0" cellspacing="1" cellpadding="3">
                    <tr valign="top">
                      <th>Image Management </th>
                      <th>Folder Management </th>
                    </tr>
                    <tr valign="top"> 
                      <td width="50%" valign="top" class="matte"> 
<form name="form1" enctype="multipart/form-data" method="post" action="<%=GP_uploadAction%>" onsubmit="checkFileUpload(this,'GIF,JPG',true,<% If MaxPhysicalSize = "" Then Response.Write "''" Else Response.Write MaxPhysicalSize %>,<% If MinWidth = "" Then Response.Write "''" Else Response.Write MinWidth %>,<% If MinHeight = "" Then Response.Write "''" Else Response.Write MinHeight %>,<% If MaxWidth = "" Then Response.Write "''" Else Response.Write MaxWidth End If%>,<% If MaxHeight = "" Then Response.Write "''" Else Response.Write MaxHeight %>,'','');showProgressWindow('fileCopyProgress.htm',300,100);return document.MM_returnValue">
                          <p>Upload to folder:<br />
                            <select name="myField">
<%
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
Dim objFolder
Set objFolder = objFSO.GetFolder(Server.MapPath(strCurrentFolder))
Dim FolderRoot, oPos
FolderRoot = objFolder
oPos = Len(FolderRoot)
ListFolders objFolder, true

Sub ListFolders(objFolder,showroot)
	Dim trimLength
	If objFolder = FolderRoot Then 
		If showroot = True Then
			Response.Write("<option value=""" & objFolder & """>")
			Response.Write "Root Image Folder"
			Response.Write("</option>" & VbCrLf)
		End If
	Else
		Response.Write("<option value=""" & objFolder & """")
		If Request.QueryString("show") = objFolder Then
			Response.Write " SELECTED"
		End If
		Response.Write(">")
		trimLength = Len(objFolder) - oPos - 1
		Response.Write(Right(objfolder,trimLength))
		Response.Write("</option>" & VbCrLf)
	End If

	'Now, use a for each...next to loop through the Files collection
	If ShowSubs = True Then
		Dim objSubFolder
		For Each objSubFolder in objFolder.SubFolders
			ListFolders objSubFolder, showroot
		Next
	End If
End Sub
%>
                            </select>
                            <br />
                            Image:<br />
                            <input type="file" name="file" onChange="checkFileUpload(this.form,'GIF,JPG',true,'','','','','','','')">
                            <br />
                            <input name="Submit" type="submit" class="go" value="Upload Image">
                          </p>
                        </form>
                      </td>
                      <td width="50%" valign="top" class="matte"> 
                        <form name="folderMgmt" method="post" action="imageupload.asp?fldrmgmt=true&img=<%= request.querystring("img") %>&path=<%= Request.QueryString("path")%>&show=<%= Server.URLEncode(Request.QueryString("show"))%>">
                          <p>Create a new folder:<br />
                            Under: 
                            <select name="lstParentFolder">
                              <% ListFolders objFolder, true %>
                            </select>
                            <br />
                            Folder Name: 
                            <input type="text" name="txtFldrName">
                            <br />
                            <input name="sbtAddFolder" type="submit" class="go" value="Add Folder">
                          </p>
                          <p>Delete a Folder:<br />
                            <select name="lstFldrName">
                              <option selected>Choose a Folder</option>
                              <% ListFolders objFolder, false %>
                            </select>
                            <br />
                            <input name="sbtDeleteFolders" type="submit" class="go" onclick="GP_popupConfirmMsg('Are you sure you want to delete this directory? All files in the directory will be deleted. This action cannot be undone.');return document.MM_returnValue" value="Delete Folder">
                          </p>
                        </form>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            
            <form name="form2" method="post" action="">
              <div align="center"> 
                <input name="Submit2" type="button" class="go" onclick="MM_goToURL('parent','images.asp?img=<%= Request.QueryString("img") %>&path=<%= Request.QueryString("path") %>&show=<%= Server.URLEncode(Request.QueryString("show")) %>');return document.MM_returnValue" value="Back to Images">
              </div>
            </form>
          </td>
        </tr>
        <tr> 
          <th colspan="2" align="center"><a href="javascript:window.close();">Close</a></th>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
