Order Manager Component Installation Guide for StoreFront 5.0

Version 1.4
Release Date: June 27, 2002

The Order Manager Component consists of the following files:

  ssAdmin (folder)
    AdminHeader.asp
    calendar.html
    clsOrderAdmin.asp
    help_OrderManager.htm
    OrderAdmin.asp
    OrderAdmin_common.asp
    OrderAdminDetail.asp
    OrderAdminSummary.asp
    OrdersImport.asp
    ShippingImportSample_comma.txt
    ShippingImportSample_tab.txt
    ShippingMethods_common.asp
    ssOrderAdminEmail.txt

  Images
    calendar.gif
    down.gif
    logo_blue.gif
    MSGBOX03.ICO
    NOTE12.ICO
    up.gif


  SSLibrary (folder)
    calendar.js
    modDatabase.asp
    ssFormValidation.js
    ssStyleSheet.css

OrderManagerSF5Upgrade.asp
ReadME.txt
ReleaseNotes.txt
SandshotLicense.htm
ssOrderManager.asp

Installation:

Step 1: Copy the files.

    When you unzipped the files it should have created the directory OrderManager.
    This directory should be placed into your root web folder.

    Place the file ssOrderManager.asp into the SFLib folder.
    Place the ssAdmin folder and its files into the ssl/Admin folder.

Step 2: Upgrade the Storefront Database.

    As always, it is recommended you make a backup of you database before running the file.

    Run the file OrderManagerSF5Upgrade.asp from you web browser.

    Summary of changes to the Storefront database:

	ssOrderManager table is created 

Step 3:  Create OrderHistory.asp

  In FrontPage select File --> New --> Page and select StoreFront 5.0 Product Page

  View the page in the HTML tab

  Delete the following lines

    <!--#include file="SFLib/product.asp"-->
    <script language="javascript" src="SFLib/sfCheckErrors.js"></script>
    <script language="javascript" src="SFLib/sfEmailFriend.js"></script>

  Find

    bot="PurpleText" PREVIEW="HEADER OR INSTRUCTIONS" -->

  and insert

    <!--#include file="SFLib/ssOrderManager.asp"-->

  immediately after it

  Optional: You may desire to change the page title

  Save the file as OrderHistory.asp in your root directory

Step 4:  (only required for SQL Server databases And early versions of StoreFront pre 50.3)

  Open the file modDatabase.asp

  Find the line 

	Const cblnSQLDatabase = False	'Set this value to True for SQL Server databases

  and change it to

	Const cblnSQLDatabase = True	'Set this value to True for SQL Server databases

Step 5: Modify ssl/sfLib/mail.asp

  Create a backup copy of mail.asp

  Open mail.asp

  Find the following line

    If sMailMethod = "Simple Mail" Then

  Immediately before it is the line


    End If

  Replace the "End If" line with the following section of code

ElseIf sType ="ssPromoMail" Then
	If Err.number = 5 Then Err.Clear
	arrInfo = split(sInformation, "|")
	sCustEmail = arrInfo(0)
	If ((len(arrInfo(1))>0) AND (arrInfo(1)<>"-")) Then sPrimary = arrInfo(1)
	If ((len(arrInfo(2))>0) AND (arrInfo(2)<>"-")) Then 
		sSecondary = arrInfo(2)
	Else
		sSecondary = "" 
	End If
	sSubject = arrInfo(3)
	sMessage = arrInfo(4)
 	sType = "PromoMail"	'setting it to PromoMail prevents the merchant from receiving a duplicate email
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

Step 6: Configure ssOrderManager.asp (Optional)

  If you want your customers to see an order history section set mblnShowOrderSummaries 
  equal to True in the user configuration section


Congratulations! You're done.