<%
'Page Design by Web Shorts Site Design, www.web-shorts.com 
'StoreFront Merchant Tools Image Management Mod 1.0
'Variables for the directory and subdirectories

'Change this to be the relative path to the image directory on the server.
Const strCurrentFolder = "../../../../images/"

'This is the absolute URL to the parent of your graphics folder. If your
'graphics folder is located at http://www.web-shorts.com/gfx/ then this
'attribute should be http://www.web-shsorts.com/ (be sure to include the
'trailing slash
Dim strAbsoluteURL: strAbsoluteURL = Application("adminDomainName")

'This is the path from your product pages to the image folder. This should
'follow the same format as strAbsoluteURL in that it should include the trailing
'slash. This can be an absolute URL, a site relative or document relative string,
'whichever suits your situation. This value will be prepended to the image path
'added into your text boxes.
Const strPathFromProducts = "../"

'Set the variable to the largest file size allowed. The number is in KBs. Leave it
'blank ("") to allow any size.
Const MaxPhysicalSize = ""

'Use these variables to set the image minimum and maximum dimentsions (height and width).
'Leave them blank ("") to allow any size image.
Const MinWidth = ""
Const MinHeight = ""
Const MaxWidth = ""
Const MaxHeight = ""
	
'Set this value to True to include subfolders. False will not include subfolders
Const ShowSubs = True

'DO NOT EDIT BELOW THIS LINE!!!!
Const vDebug = 0

Function wl(byVal Text)
	Response.Write Text & VbCrLf
End Function
Function DebugValues
	wl("<h4>Debugging, some things may look weird or function incorrectly.</h4>")
	wl("<p><b>Querystring Values</b></p>")
	wl("<table border=""1"" cellspacing=""0"" cellpadding=""0"">")
	wl("<tr>")
	wl("<td>Variable</td>")
	wl("<td>Value</td>")
	wl("</tr>")
	Dim Item
	For Each Item In Request.Querystring
		wl("<tr>")
		wl("<td>" & Item & "</td>")
		wl("<td>" & Request.Querystring(Item) & "</td>")
		wl("</tr>")
	Next
	wl("</table>")
End Function

'*** Pure ASP File Upload -----------------------------------------------------
' Copyright 2001 (c) George Petrov, www.UDzone.com
'
' Script partially based on code from Philippe Collignon 
'              (http://www.asptoday.com/articles/20000316.htm)
'
' New features from GP:
'  * Fast file save with ADO 2.5 stream object
'  * new file handling, wrapper functions, extra error checking
'  * UltraDev Server Behavior extension
'
' Version: 2.0.3
'------------------------------------------------------------------------------
Sub BuildUploadRequest(RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict)
  'Get the boundary
  PosBeg = 1
  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
  if PosEnd = 0 then
    Response.Write "<b>Form was submitted with no ENCTYPE=""multipart/form-data""</b><br />"
    Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"    
    Response.End
  end if
  'Check ADO Version
	set checkADOConn = Server.CreateObject("ADODB.Connection")
	adoVersion = CSng(checkADOConn.Version)
	set checkADOConn = Nothing
	if adoVersion < 2.5 then
    Response.Write "<b>You don't have ADO 2.5 installed on the server.</b><br />"
    Response.Write "The File Upload extension needs ADO 2.5 or greater to run properly.<br />"
    Response.Write "You can download the latest MDAC (ADO is included) from <a href=""www.microsoft.com/data"">www.microsoft.com/data</a><br />"
    Response.End
	end if		
  'Check content length if needed
	Length = CLng(Request.ServerVariables("HTTP_Content_Length")) 'Get Content-Length header
	If "" & sizeLimit <> "" Then
    sizeLimit = CLng(sizeLimit) * 1024
    If Length > sizeLimit Then
      Request.BinaryRead (Length)
      Response.Write "Upload size " & FormatNumber(Length, 0) & "B exceeds limit of " & FormatNumber(sizeLimit, 0) & "B. "
      Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"      
      Response.End
    End If
  End If
  boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
  boundaryPos = InstrB(1,RequestBin,boundary)
  'Get all data inside the boundaries
  Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
    'Members variable of objects are put in a dictionary object
    Dim UploadControl
    Set UploadControl = CreateObject("Scripting.Dictionary")
    'Get an object name
    Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
    Pos = InstrB(Pos,RequestBin,getByteString("name="))
    PosBeg = Pos+6
    PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
    Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
    PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)
    'Test if object is of file type
    If  PosFile<>0 AND (PosFile<PosBound) Then
      'Get Filename, content-type and content of file
      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
      FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      FileName = Mid(FileName,InStrRev(FileName,"\")+1)
      'Add filename to dictionary object
      UploadControl.Add "FileName", FileName
      Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
      PosBeg = Pos+14
      PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
      'Add content-type to dictionary object
      ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      UploadControl.Add "ContentType",ContentType
      'Get content of object
      PosBeg = PosEnd+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = FileName
      ValueBeg = PosBeg-1
      ValueLen = PosEnd-Posbeg
    Else
      'Get content of object
      Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
      PosBeg = Pos+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      ValueBeg = 0
      ValueEnd = 0
    End If
    'Add content to dictionary object
    UploadControl.Add "Value" , Value	
    UploadControl.Add "ValueBeg" , ValueBeg
    UploadControl.Add "ValueLen" , ValueLen	
    'Add dictionary object to main dictionary
    UploadRequest.Add name, UploadControl	
    'Loop to next object
    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
  Loop

  GP_keys = UploadRequest.Keys
  for GP_i = 0 to UploadRequest.Count - 1
    GP_curKey = GP_keys(GP_i)
    'Save all uploaded files
    if UploadRequest.Item(GP_curKey).Item("FileName") <> "" then
      GP_value = UploadRequest.Item(GP_curKey).Item("Value")
      GP_valueBeg = UploadRequest.Item(GP_curKey).Item("ValueBeg")
      GP_valueLen = UploadRequest.Item(GP_curKey).Item("ValueLen")

      'Get the path
      'if InStr(UploadDirectory,"\") then
      '  GP_curPath = UploadDirectory
      '  if Mid(GP_curPath,Len(GP_curPath),1)  <> "\" then
      '    GP_curPath = GP_curPath & "\"
      '  end if         
      '  GP_FullPath = GP_curPath
      'else
        'GP_curPath = Request.ServerVariables("PATH_INFO")
        'GP_curPath = Trim(Mid(GP_curPath,1,InStrRev(GP_curPath,"/")) & UploadDirectory)
        'if Mid(GP_curPath,Len(GP_curPath),1)  <> "/" then
        '  GP_curPath = GP_curPath & "/"
        'end if 
        'GP_FullPath = Trim(Server.mappath(GP_curPath))
      '  GP_FullPath = Trim(Server.mappath(UploadDirectory))
      'end if
		GP_FullPath = UploadFormRequest("myField") & "\"
      
      if GP_valueLen = 0 then
        Response.Write "<B>An error has occured saving uploaded file!</B><br /><br />"
        Response.Write "Filename: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br />"
        Response.Write "File does not exists or is empty.<br />"
        Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
	  	  response.End
	    end if
      
      'Create a Stream instance
      Dim GP_strm1, GP_strm2
      Set GP_strm1 = Server.CreateObject("ADODB.Stream")
      Set GP_strm2 = Server.CreateObject("ADODB.Stream")
      
      'Open the stream
      GP_strm1.Open
      GP_strm1.Type = 1 'Binary
      GP_strm2.Open
      GP_strm2.Type = 1 'Binary
        
      GP_strm1.Write RequestBin
      GP_strm1.Position = GP_ValueBeg
      GP_strm1.CopyTo GP_strm2,GP_ValueLen
    
      'Create and Write to a File
      GP_CurFileName = UploadRequest.Item(GP_curKey).Item("FileName")      
      GP_FullFileName = GP_FullPath & "\" & GP_CurFileName
      Set fso = CreateObject("Scripting.FileSystemObject")
      'Check if the folder exist
      If NOT fso.FolderExists(GP_FullPath) Then
        GP_BegFolder = InStr(GP_FullPath,"\")
        while GP_begFolder > 0 
          GP_RelFolder = Mid(GP_FullPath,1,GP_BegFolder-1)
          If NOT fso.FolderExists(GP_RelFolder) Then  
            fso.CreateFolder(GP_RelFolder)
          end if          
          GP_BegFolder = InStr(GP_BegFolder+1,GP_FullPath,"\")          
        wend
        If NOT fso.FolderExists(GP_FullPath) Then        
          fso.CreateFolder(GP_FullPath)        
        end if  
      end if
      'Check if the file already exist
      GP_FileExist = false
      If fso.FileExists(GP_FullFileName) Then
        GP_FileExist = true
      End If      
      if nameConflict = "error" and GP_FileExist then
        Response.Write "<B>File already exists!</B><br /><br />"
        Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
				GP_strm1.Close
				GP_strm2.Close
	  	  response.End
      end if
      if ((nameConflict = "over" or nameConflict = "uniq") and GP_FileExist) or (NOT GP_FileExist) then
        if nameConflict = "uniq" and GP_FileExist then
          Begin_Name_Num = 0
          while GP_FileExist    
            Begin_Name_Num = Begin_Name_Num + 1
            GP_FullFileName = Trim(GP_FullPath)& "\" & fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
            GP_FileExist = fso.FileExists(GP_FullFileName)
          wend  
          UploadRequest.Item(GP_curKey).Item("FileName") = fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
					UploadRequest.Item(GP_curKey).Item("Value") = UploadRequest.Item(GP_curKey).Item("FileName")
        end if
        on error resume next
        GP_strm2.SaveToFile GP_FullFileName,2
        if err then
          Response.Write "<B>An error has occured saving uploaded file!</B><br /><br />"
          Response.Write "Filename: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br />"
          Response.Write "Maybe the destination directory does not exist, or you don't have write permission.<br />"
          Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
    		  err.clear
  				GP_strm1.Close
  				GP_strm2.Close
  	  	  response.End
  	    end if
  			GP_strm1.Close
  			GP_strm2.Close
  			if storeType = "path" then
  				UploadRequest.Item(GP_curKey).Item("Value") = GP_curPath & UploadRequest.Item(GP_curKey).Item("Value")
  			end if
        on error goto 0
      end if
    end if
  next

End Sub

'String to byte string conversion
Function getByteString(StringStr)
  For i = 1 to Len(StringStr)
 	  char = Mid(StringStr,i,1)
	  getByteString = getByteString & chrB(AscB(char))
  Next
End Function

'Byte string to string conversion
Function getString(StringBin)
  getString =""
  For intCount = 1 to LenB(StringBin)
	  getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
  Next
End Function

Function UploadFormRequest(name)
  on error resume next
  if UploadRequest.Item(name) then
    UploadFormRequest = UploadRequest.Item(name).Item("Value")
  end if  
End Function

Sub PureUploadSetup()
  UploadQueryString = Replace(Request.QueryString,"GP_upload=true","")
  if mid(UploadQueryString,1,1) = "&" then
  	UploadQueryString = Mid(UploadQueryString,2)
  end if
  GP_uploadAction = CStr(Request.ServerVariables("URL")) & "?GP_upload=true"
  If (Request.QueryString <> "") Then  
    if UploadQueryString <> "" then
  	  GP_uploadAction = GP_uploadAction & "&" & UploadQueryString
    end if 
  End If
End Sub
%>
