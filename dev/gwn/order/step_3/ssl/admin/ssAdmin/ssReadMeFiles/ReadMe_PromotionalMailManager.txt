Promotional Mail Manager Component Installation Guide for StoreFront 5.0

Version 1.0
Release Date: October 28, 2001

The Order Manager Component consists of the following files:

  ssAdmin (folder)
    AdminHeader.asp
    calendar.html
    help_PromotionalMailManager_help.htm
    PromoMailAdmin.asp
    PromoMailSample.htm
    PromoMailSample.txt

  Images
    calendar.gif
    down.gif
    logo_blue.gif
    up.gif


  SSLibrary (folder)
    calendar.js
    modDatabase.asp
    ssFormValidation.js

OrderManagerSF5Upgrade.asp
ReadME.txt
SandshotLicense.htm
ssOrderManager.asp

Installation:

Step 1: Copy the files.

    When you unzipped the files it should have created the directory PromoMailManager.
    This directory should be placed into your root web folder.

     Place the ssAdmin folder and its files into the ssl/Admin folder.

Step 2: Modify ssl/sfLib/mail.asp

  Create a backup copy of mail.asp

  Open mail.asp

  Find the following line

    If sMailMethod = "Simple Mail" Then

  Immediately before it is the line


    End If

  Replace the "End If" line with the following section of code

    Else
	If Err.number = 5 Then Err.Clear
	arrInfo = split(sInformation, "|")
	sCustEmail = arrInfo(0)
	If ((len(arrInfo(1))>0) AND (arrInfo(1)<>"-")) Then sPrimary = arrInfo(1)
	If ((len(arrInfo(2))>0) AND (arrInfo(2)<>"-")) Then sSecondary = arrInfo(2)
	sSubject = arrInfo(3)
	sMessage = arrInfo(4)
 	sType = "PromoMail"	'setting it to PromoMail prevents the merchant from receiving a duplicate email
   End If


Congratulations! You're done.