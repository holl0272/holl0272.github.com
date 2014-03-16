<%

Dim strContent

strContent = Request.Form("Content")

%>
<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>

	<script language="JavaScript" src="include/yusasp_ace.js"></script>
	<script language="JavaScript" src="include/yusasp_color.js"></script>
	
	<script>
	function SubmitForm()
		{
		if(objACE.displayMode == "HTML")
		{
			alert("You must uncheck HTML view in order to save.")
			return ;
		}

		parent.opener.ACEfield.value = objACE.getContentBody()
		window.close();
		}
		
	function LoadContent()
	{
		document.all("fieldBeingEdited").innerText = "Field being edited: " + parent.opener.ACEfield.title;
		document.title = "Field being edited: " + parent.opener.ACEfield.title;
		objACE.putContent(parent.opener.ACEfield.value) 
	}
	</script>
</head>

<body onload="LoadContent()" style="font:10pt verdana,arial,sans-serif">

<h4 id="fieldBeingEdited">Field being edited: Loading data . . .</h4>
<script>
	var objACE = new ACEditor("objACE")
	objACE.width = "100%" 
	objACE.height = 300
	objACE.useStyle = false
	objACE.useAsset = false
	objACE.useImage = false
	objACE.usePageProperties = false
	objACE.RUN() 
</script>

<input type="button" NAME="btnCancel" ID="btnCancel" onclick="window.close();" value="Cancel">
<input type="button" NAME="btnSave" ID="btnSave" onclick="SubmitForm()" value="S A V E">

</body>
</html>
