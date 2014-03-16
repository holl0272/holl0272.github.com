//@BEGINVERSIONINFO

//@APPVERSION: 50.4014.0.3

//@FILENAME: sfcheckerrors.asp 
	 


//@DESCRIPTION: Checkes sfErrors

//@STARTCOPYRIGHT
//The contents of this file is protected under the United States
//copyright laws as an unpublished work, and is confidential and proprietary to
//LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
//expressed written permission of LaGarde, Incorporated is expressly prohibited.

//(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
//@ENDCOPYRIGHT

//@ENDVERSIONINFO

function specialCase(e, form) {
	if ((e.name == "CardName")||(e.name == "CardNumber")||(e.name == "CardExpiryMonth")||(e.name == "CardExpiryYear")) {
		if (((form.CardName.value.length <= 0)||(form.CardNumber.value.length <= 0)||(form.CardExpiryMonth.value.length <= 0)||(form.CardExpiryYear.value.length <= 0))
		 && ((form.CardName.value.length > 0)||(form.CardNumber.value.length > 0)||(form.CardExpiryMonth.value.length > 0)||(form.CardExpiryYear.value.length > 0))) {
			return "Please enter all Credit Card Information.";
		}
		if ((form.CardName.value.length > 0)&&(form.CardNumber.value.length > 0)&&(form.CardExpiryMonth.value.length > 0)&&(form.CardExpiryYear.value.length > 0)) {
			if (!isCardDateValid(form.CardExpiryYear.value, form.CardExpiryMonth.value)) {
				return "The Credit Card has Expired.";
			}
			if (isCardNumValid(form.CardNumber.value))
			{
				return "The Credit Card Number is an invalid format.";
			}else{
				if (!isCorrectCreditCardType(form.CardType, form.CardNumber))
				{
					return "The Credit Card Number is invalid for the selected Card Type.";
				}
			}
		}
	}
	if ((e.name == "CheckNumber")||(e.name == "BankName")||(e.name == "RoutingNumber")||(e.name == "CheckingAccountNumber")) {
		if (((form.CheckNumber.value.length <= 0)||(form.BankName.value.length <= 0)||(form.RoutingNumber.value.length <= 0)||(form.CheckingAccountNumber.value.length <= 0))
		 && ((form.CheckNumber.value.length > 0)||(form.BankName.value.length > 0)||(form.RoutingNumber.value.length > 0)||(form.CheckingAccountNumber.value.length > 0))) {
			return "Please enter all eCheck Information.";
		}		
	}
	if ((e.name == "POName")||(e.name == "PONumber")) {
		
		if (((form.POName.value.length <= 0)||(form.PONumber.value.length <= 0))
		 && ((form.POName.value.length > 0)||(form.PONumber.value.length > 0))) {
			return "Please enter all Purchase Order Information.";
		}
	}
	if ((form.CardName.value.length <= 0)&&(form.CardNumber.value.length <= 0)&&(form.CardExpiryMonth.value.length <= 0)&&(form.CardExpiryYear.value.length <= 0)
	 && (form.CheckNumber.value.length <= 0)&&(form.BankName.value.length <= 0)&&(form.RoutingNumber.value.length <= 0)&&(form.CheckingAccountNumber.value.length <= 0)
	 && (form.POName.value.length <= 0)&&(form.PONumber.value.length <= 0)) {
		return "Please enter payment method Information.";	
	}
	return "";
}
function stripChar(sValue, sChar) {
	var i, tempChar, buildString;
	buildString = ""
	for (var i=0; i<sValue.length; i++) {
		tempChar = sValue.charAt(i);
		if (tempChar != sChar) {
			buildString = buildString + tempChar;
		}
	}
	return buildString;
}

function isCardDateValid(year, month) {
	var dateCheck, now;
	if (year.length == 2) {
		if (parseInt(year) < 50) {
			year = "20" + year;
		}
	}
	now = new Date();
	dateCheck = new Date(year, month);
	if (now > dateCheck) {
		return false;
	}
	else {
		return true;
	}
}

function isCardNumValid(num) {
	var num1, num2, tempNum;
	if (!isNumber(num)) {
		return true;
	}
	num1 = ""
	if (!(num.length%2==0)) {
		for(var j=0; j < num.length; j++) {
			if ((j+1)%2==0){
				tempNum = 2 * num.charAt(j);
			}
			else {
				tempNum = 1 * num.charAt(j);
			}
			num1 = num1 + tempNum.toString();
		}
	}
	else{
		for(var j=0; j < num.length; j++){
			if ((j+1)%2==0){
				tempNum = 1 * num.charAt(j);
			}
			else{
				tempNum = 2 * num.charAt(j);
			}
			num1 = num1 + tempNum.toString();
		}
	}
	num2 = 0;
	for (var j = 0; j < num1.length; j++) {
		num2 = num2 + parseInt(num1.charAt(j));
	}
	if (num2%10==0) {
		return false;
	}
	else {
		return true;
	}
}

function isNumber(value) {
	for (var i=0; i < value.length; i++) {
		a = parseInt(value.charAt(i));
		if (isNaN(a)) {
			return false;			
			break;
		}
	}
	return true;
}

function sfCheck(form) {
	var e, title, empty_fields, char_check, invalid_card, month, year, invalid_date, eMail, invalid_eMail 
	var iQuantity, quantity_check, checkSpecial, tempError, special_Error, msg, upperLine, lowerLine
	var bad_Zip,num, invalid_phoneNumber, passwd_mismatch
	msg = "";

	empty_fields = "";
	char_check = "";
	special_Error = "";
	tempError = "";
	num = form.length
	for (var i = 0; i < form.length; i++) {
		e = form.elements[i]
		if ((e.title == null)||(e.title == "")) {
			title = e.name;
		}
		else {
			title = e.title;
		}
		if (((e.type == "text") || (e.type == "textarea")||(e.type == "password")) && !e.special && !e.disabled) {
			if (e.value.length <= 0 && !e.optional && (e.name.indexOf("Ship") == -1)) {
				empty_fields += "\n            " + title;
				continue;
			}
			if (e.number) {
				num = e.value;
				num = stripChar(num, ".");
				num = stripChar(num, ",");
				if (!isNumber(num)) {
					char_check += "\n             " + title;
				}
			}
			if (e.creditCardNumber) {
				e.value = stripChar(e.value, " "); 
				e.value = stripChar(e.value, "-"); 
				invalid_card = isCardNumValid(e.value);
				
				if (!isCardNumValid(form.CardNumber.value))
				{
					if (!isCorrectCreditCardType(form.CardType, form.CardNumber))
					{
						alert("The Credit Card Number is invalid for the selected Card Type.");
						return false;
					}
				}
			}
			
			if ((e.creditCardExpMonth)||(e.creditCardExpYear)) {
				if (e.creditCardExpMonth) {
					month = e.value;
					month = stripChar(month, " ")
					if (!isNumber(month)) {
						invalid_date = true;
						month = null;
					}
				}
				if (e.creditCardExpYear) {
					year = e.value;
					year = stripChar(year, " ")
					if (!isNumber(year)) {
						invalid_date = true;
						year = null;
					}
				}
				if ((month != null) && (year != null)) {
					if(!isCardDateValid(year, month)) {
						invalid_date = true;
					}	
				}
			}
			if (e.eMail) {
				eMail = e.value;
				if ((eMail.indexOf("@") != -1) && (eMail.indexOf(".") != -1)) {
					invalid_eMail = false;
				}
				else {
					invalid_eMail = true;
				}
			}
            if (e.name == "txtEmail") {
				eMail = e.value;
				if ((eMail.indexOf("@") != -1) && (eMail.indexOf(".") != -1)) {
					invalid_eMail = false;
				}
				else {
				  
					invalid_eMail = true;
				}
			}	
            if (e.name == "Email") {
				eMail = e.value;
				if ((eMail.indexOf("@") != -1) && (eMail.indexOf(".") != -1)) {
					invalid_eMail = false;
				}
				else {
				  
					invalid_eMail = true;
				}
			}	
			if (e.name == "txtFriend") {
				eMail = e.value;
				if ((eMail.indexOf("@") != -1) && (eMail.indexOf(".") != -1)) {
					invalid_eMail = false;
				}
				else {
				  
					invalid_eMail = true;
				}
			}	

			if (e.phoneNumber) {
				num = e.value;
				num = stripChar(num, " ");
				num = stripChar(num, "-");
				num = stripChar(num, "+");
				if (num.length < 10) {
					invalid_phoneNumber = true;
				}	
			}
		}
		if (e.quantityBox) {
			iQuantity = e.value;
			if (!isNumber(iQuantity)) {
				quantity_check = true;
			}
			if (parseInt(iQuantity) < 0) {
				quantity_check = true;
			}
			if ((iQuantity) < 1) {
				quantity_check = true;
			}

		}
		if (e.password) {
			if (form.Password.value != form.Password2.value) {
					passwd_mismatch = true;
			}
		}
		if (e.zipcode) {
			if (e.value.length<minZipLength(e)) {

				bad_Zip = true;
			}
		}

		if (e.special) {
			checkSpecial = specialCase(e, form);
			if (tempError != checkSpecial) {
				special_Error = special_Error + checkSpecial
			}
			tempError = checkSpecial;
		}
		if (e.type == "select-one" && !e.optional) {
			if (e.options.selectedIndex != -1)
			{
				if (e.options[e.options.selectedIndex].value == "") {
					empty_fields += "\n            " + title;
					continue;
				}
			}else{
				empty_fields += "\n            " + title;
				continue;
			}
		}
	}
	
	if (!bad_Zip && !empty_fields && !char_check && !special_Error && !invalid_card && !invalid_date && !invalid_eMail && !quantity_check && !invalid_phoneNumber && !passwd_mismatch) {return true}
	
	msg = "The form was not submited due to the following error(s).\n";
	
	upperLine = "\n_________________________________________________________\n\n";
	lowerLine = "_________________________________________________________\n";
	
	if (empty_fields) {
		msg += upperLine;
		msg += "The following field(s) must be filled in:\n";
		msg += lowerLine;
		msg += empty_fields;
	}
	if (char_check) {
		msg += upperLine;
		msg += "The following field(s) need a numeric value:\n";
		msg += lowerLine;
		msg += char_check;
	}
	if (quantity_check) {
		msg += upperLine;
		msg += "Please Enter a Positive Integer.\n"
		msg += lowerLine;
	}
	if (invalid_card) {
		msg += upperLine;
		msg += "The Credit Card Number is an invalid format.\n";
		msg += lowerLine;
	}
	if (invalid_date) {
		msg += upperLine;
		msg += "The Credit Card has Expired.\n";
		msg += lowerLine;
	}
	if (invalid_eMail) {
		msg += upperLine;
		msg += "The Email Address is in an invalid format.\n";
		msg += lowerLine;
	}
	if (invalid_phoneNumber) {
		msg += upperLine;
		msg += "Please enter a valid Phone Number with area code.\n";
		msg += lowerLine;
	}
	if (special_Error) {
		msg += upperLine;
		msg += special_Error + "\n";
		msg += lowerLine;
	}
	if (passwd_mismatch) {
		msg += upperLine;
		msg += "Your passwords did not match. Please enter them again.\n";
		msg += lowerLine;
	}	
	if (bad_Zip) {
		msg += upperLine;
		msg += "The postal code is too short. Please enter it again.\n";
		msg += lowerLine;
	}		
	
	alert(msg);
	return false;
}	

function sfCheckPlus(frm) {
  if (window.document.form1.ShipZip.value != "" || window.document.form1.ShipFirstName.value != "" || window.document.form1.ShipMiddleInitial.value != "" || window.document.form1.ShipLastName.value != "" || window.document.form1.ShipCompany.value != "" || window.document.form1.ShipAddress1.value != "" || window.document.form1.ShipAddress2.value != "" || window.document.form1.ShipCity.value != "" || (window.document.form1.ShipState.value != "" && window.document.form1.ShipState.value != null) || (window.document.form1.ShipCountry.value != "" && window.document.form1.ShipCountry.value != null) || window.document.form1.ShipPhone.value != "" || window.document.form1.ShipEmail.value != "" || window.document.form1.ShipFax.value != "")			  
	{
	var blnShipStateRequired;
	
	blnShipStateRequired = (window.document.form1.ShipCountry.options[document.form1.ShipCountry.selectedIndex].text  == "us" || window.document.form1.ShipCountry.options[document.form1.ShipCountry.selectedIndex].text  == "ca");

	//if ((window.document.form1.ShipZip.value == "" && window.document.form1.ShipZip.optional==false) || window.document.form1.ShipFirstName.value == "" || window.document.form1.ShipLastName.value == "" || window.document.form1.ShipAddress1.value == "" || window.document.form1.ShipCity.value == "" || document.form1.ShipState.options[document.form1.ShipState.selectedIndex].text  ==""  || window.document.form1.ShipCountry.options[document.form1.ShipCountry.selectedIndex].text  == "" || window.document.form1.ShipPhone.value == "" || window.document.form1.ShipEmail.value == "")			  
	if ((window.document.form1.ShipZip.value == "" && window.document.form1.ShipZip.optional==false) || window.document.form1.ShipFirstName.value == "" || window.document.form1.ShipLastName.value == "" || window.document.form1.ShipAddress1.value == "" || window.document.form1.ShipCity.value == "" || (document.form1.ShipState.options[document.form1.ShipState.selectedIndex].text  == "" && blnShipStateRequired) || window.document.form1.ShipCountry.options[document.form1.ShipCountry.selectedIndex].text  == "" || window.document.form1.ShipPhone.value == "" || window.document.form1.ShipEmail.value == "")			  
	{
	window.alert("Please either fill in all shipping info or no shipping info.")
	return false
	}
  else
	{return sfCheck(frm)}
  }		  

}
function POCheck(poname,poNum)
{
  if(poname == "" || poNum == "")
   {
    alert("Please Enter the required purchase order information"); 
      
    return false;
   } 
   else
    {
     return true;
    }
}

function ECheck(frm)
{
	if (frm.CheckNumber.value == "" || frm.BankName.value == "" ||
		frm.RoutingNumber.value == "" || frm.CheckingAccountNumber.value == "")
	{
    alert("Please Enter the required e-check information"); 
      
    return false;
	}
	else
	{
		return true;
	}
}

function isInteger(theField, emptyOK, theMessage)
{
	if (theField.value == "")
	{
	if (emptyOK)
	{
		return(true);
	}
	{
		alert(theMessage);
		theField.focus();
		theField.select();
		return (false);
	}
	}

	var i;
	var s = theField.value;
	for (i = 0; i < s.length; i++)
	{
	    var c = s.charAt(i);
	    if (!((c >= "0") && (c <= "9")))
	    {
			alert(theMessage);
			theField.focus();
			theField.select();
	        return (false);
	    }
	}
	if (s > 0)
	{
	return (true);
	}
	{
	alert(theMessage);
	theField.focus();
	theField.select();
	return (false);
	}
}


function isCorrectCreditCardType (theSelect, theField)
{
	var cardType = theSelect.options[theSelect.selectedIndex].value;
    var normalizedCCN = theField.value;
    
    normalizedCCN = stripChar(normalizedCCN, " "); 
	normalizedCCN = stripChar(normalizedCCN, "-"); 

	//alert(cardType + ' - ' + normalizedCCN);
    if (isCardMatch(cardType, normalizedCCN)) return true;
    return false;
}

function isAnyCard(cc)
{
  if (!isMasterCard(cc) && !isVisa(cc) && !isAmericanExpress(cc) && !isDinersClub(cc) && !isDiscover(cc) && !isEnRoute(cc) && !isJCB(cc)) return false;
  return true;
}

function isCardMatch (cardType, cardNumber)
{
	if (cardType == "1") return (isAmericanExpress(cardNumber));
	if (cardType == "2") return (isVisa(cardNumber));
	if (cardType == "3") return (isDiscover(cardNumber));
	if (cardType == "4") return (isMasterCard(cardNumber));
	if (cardType == "5") return (isDinersClub(cardNumber));
	if (cardType == "6") return (isCarteBlanche(cardNumber));

	if (cardType == "xx") return (isEnRoute(cardNumber));
	if (cardType == "xx") return (isJCB(cardNumber));

	return false;

}

function isAmericanExpress(cc)
{
  firstdig = cc.substring(0,1);
  seconddig = cc.substring(1,2);
  if ((cc.length == 15) && (firstdig == 3) && ((seconddig == 4) || (seconddig == 7))) return true;
  return false;
}

function isCarteBlanche(cc)
{
  return isDinersClub(cc);
}

function isDinersClub(cc)
{
  firstdig = cc.substring(0,1);
  seconddig = cc.substring(1,2);
  if ((cc.length == 14) && (firstdig == 3) && ((seconddig == 0) || (seconddig == 6) || (seconddig == 8))) return true;
  return false;
}

function isDiscover(cc)
{
  first4digs = cc.substring(0,4);
  if ((cc.length == 16) && (first4digs == "6011")) return true;
  return false;
}

function isEnRoute(cc)
{
  first4digs = cc.substring(0,4);
  if ((cc.length == 15) && ((first4digs == "2014") || (first4digs == "2149"))) return true;
  return false;
}

function isJCB(cc)
{
  first4digs = cc.substring(0,4);
  if ((cc.length == 16) && ((first4digs == "3088") || (first4digs == "3096") || (first4digs == "3112") || (first4digs == "3158") || (first4digs == "3337") || (first4digs == "3528"))) return true;
  return false;
}

function isMasterCard(cc)
{
  firstdig = cc.substring(0,1);
  seconddig = cc.substring(1,2);
  if ((cc.length == 16) && (firstdig == 5) && ((seconddig >= 1) && (seconddig <= 5))) return true;
  return false;
}

function isVisa(cc)
{
  if (((cc.length == 16) || (cc.length == 13)) && (cc.substring(0,1) == 4)) return true;
  return false;
}

function isValidEmailAddress(email)
{
	if ((email.indexOf("@") == -1) || (email.indexOf(".") == -1))
	{
		return false;
	}

	return true;
}





<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- V1.1.3: Sandeep V. Tamhankar (stamhankar@hotmail.com) -->
<!-- Original:  Sandeep V. Tamhankar (stamhankar@hotmail.com) -->
<!-- Changes:
/* 1.1.4: Fixed a bug where upper ASCII characters (i.e. accented letters
international characters) were allowed.

1.1.3: Added the restriction to only accept addresses ending in two
letters (interpreted to be a country code) or one of the known
TLDs (com, net, org, edu, int, mil, gov, arpa), including the
new ones (biz, aero, name, coop, info, pro, museum).  One can
easily update the list (if ICANN adds even more TLDs in the
future) by updating the knownDomsPat variable near the
top of the function.  Also, I added a variable at the top
of the function that determines whether or not TLDs should be
checked at all.  This is good if you are using this function
internally (i.e. intranet site) where hostnames don't have to 
conform to W3C standards and thus internal organization e-mail
addresses don't have to either.
Changed some of the logic so that the function will work properly
with Netscape 6.

1.1.2: Fixed a bug where trailing . in e-mail address was passing
(the bug is actually in the weak regexp engine of the browser; I
simplified the regexps to make it work).

1.1.1: Removed restriction that countries must be preceded by a domain,
so abc@host.uk is now legal.  However, there's still the 
restriction that an address must end in a two or three letter
word.

1.1: Rewrote most of the function to conform more closely to RFC 822.

1.0: Original  */
// -->

function emailCheck (emailStr) {

	/* The following variable tells the rest of the function whether or not
	to verify that the address ends in a two-letter country or well-known
	TLD.  1 means check it, 0 means don't. */

	var checkTLD=1;

	/* The following is the list of known TLDs that an e-mail address must end with. */

	var knownDomsPat=/^(com|net|org|edu|int|mil|gov|arpa|biz|aero|name|coop|info|pro|museum)$/;

	/* The following pattern is used to check if the entered e-mail address
	fits the user@domain format.  It also is used to separate the username
	from the domain. */

	var emailPat=/^(.+)@(.+)$/;

	/* The following string represents the pattern for matching all special
	characters.  We don't want to allow special characters in the address. 
	These characters include ( ) < > @ , ; : \ " . [ ] */

	var specialChars="\\(\\)><@,;:\\\\\\\"\\.\\[\\]";

	/* The following string represents the range of characters allowed in a 
	username or domainname.  It really states which chars aren't allowed.*/

	var validChars="\[^\\s" + specialChars + "\]";

	/* The following pattern applies if the "user" is a quoted string (in
	which case, there are no rules about which characters are allowed
	and which aren't; anything goes).  E.g. "jiminy cricket"@disney.com
	is a legal e-mail address. */

	var quotedUser="(\"[^\"]*\")";

	/* The following pattern applies for domains that are IP addresses,
	rather than symbolic names.  E.g. joe@[123.124.233.4] is a legal
	e-mail address. NOTE: The square brackets are required. */

	var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/;

	/* The following string represents an atom (basically a series of non-special characters.) */

	var atom=validChars + '+';

	/* The following string represents one word in the typical username.
	For example, in john.doe@somewhere.com, john and doe are words.
	Basically, a word is either an atom or quoted string. */

	var word="(" + atom + "|" + quotedUser + ")";

	// The following pattern describes the structure of the user

	var userPat=new RegExp("^" + word + "(\\." + word + ")*$");

	/* The following pattern describes the structure of a normal symbolic
	domain, as opposed to ipDomainPat, shown above. */

	var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$");

	/* Finally, let's start trying to figure out if the supplied address is valid. */

	/* Begin with the coarse pattern to simply break up user@domain into
	different pieces that are easy to analyze. */

	var matchArray=emailStr.match(emailPat);

	if (matchArray==null)
	{

		/* Too many/few @'s or something; basically, this address doesn't
		even fit the general mould of a valid e-mail address. */

		alert("Email address seems incorrect (check @ and .'s)");
		return false;
	}
	
	var user=matchArray[1];
	var domain=matchArray[2];

	// Start by checking that only basic ASCII characters are in the strings (0-127).

	for (i=0; i<user.length; i++) {
	if (user.charCodeAt(i)>127) {
	alert("Ths username contains invalid characters.");
	return false;
	}
	}
	
	for (i=0; i<domain.length; i++) {
	if (domain.charCodeAt(i)>127) {
	alert("Ths domain name contains invalid characters.");
	return false;
	}
	}

	// See if "user" is valid 

	if (user.match(userPat)==null) {

	// user is not valid

	alert("The username doesn't seem to be valid.");
	return false;
	}

	/* if the e-mail address is at an IP address (as opposed to a symbolic
	host name) make sure the IP address is valid. */

	var IPArray=domain.match(ipDomainPat);
	if (IPArray!=null) {

	// this is an IP address

	for (var i=1;i<=4;i++) {
	if (IPArray[i]>255) {
	alert("Destination IP address is invalid!");
	return false;
	}
	}
	
	return true;
	}

	// Domain is symbolic name.  Check if it's valid.
	 
	var atomPat=new RegExp("^" + atom + "$");
	var domArr=domain.split(".");
	var len=domArr.length;
	for (i=0;i<len;i++) {
	if (domArr[i].search(atomPat)==-1) {
	alert("The domain name does not seem to be valid.");
	return false;
	}
	}

	/* domain name seems valid, but now make sure that it ends in a
	known top-level domain (like com, edu, gov) or a two-letter word,
	representing country (uk, nl), and that there's a hostname preceding 
	the domain or country. */

	if (checkTLD && domArr[domArr.length-1].length!=2 && 
	domArr[domArr.length-1].search(knownDomsPat)==-1) {
	alert("The address must end in a well-known domain or two letter " + "country.");
	return false;
	}

	// Make sure there's a host name preceding the domain.

	if (len<2) {
	alert("This address is missing a hostname!");
	return false;
	}

// If we've gotten this far, everything's valid!
return true;
}