
checkSessionCookie();

function checkSessionCookie()
{

var ret=ReadCookie("SFCookieEnabled");

	if (ret!="True")
	{
		SetCookie("SFCookieEnabled","True");
		ret=ReadCookie("SFCookieEnabled");
	}
	
	//ret="False";
	if (ret!="True")
	{
		window.alert("Cookies must be enabled to continue.");
		//window.history.back();
		window.location = "noCookies.asp";
	}
	

}

function ReadCookie(cookieName)
{
	var theCookie=""+document.cookie;
	var ind=theCookie.indexOf(cookieName);
	if (ind==-1 || cookieName=="") return "";
	var ind1=theCookie.indexOf(';',ind);
	if (ind1==-1) ind1=theCookie.length; 
	return unescape(theCookie.substring(ind+cookieName.length+1,ind1));
}

function SetCookie(cookieName,cookieValue,nDays)
{
	var today = new Date();
	var expire = new Date();
	if (nDays==null || nDays==0) nDays=1;
	expire.setTime(today.getTime() + 3600000*24*nDays);
	document.cookie = cookieName+"="+escape(cookieValue) + ";expires="+expire.toGMTString();
}
