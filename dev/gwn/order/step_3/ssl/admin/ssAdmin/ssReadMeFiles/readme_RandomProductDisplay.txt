Random Product Display Instructions for StoreFront 5.0

Version: 1.0
Release Date: August 26, 2001

Files:

readme_RandomProductDisplay.txt
ssRandomProductExample.asp
ssRandomProducts.asp

1) Place a copy of ssRandomProducts.asp in your sfLib folder

2) Configure ssRandomProducts.asp

'////////////////////////////////////////////////////////////////////////////////
'//
'//		USER CONFIGURATION


		pblnIsSQL = False
		pblnIsSF5AE = False
		pblnDisplayOnlyProductsWithImages = True
		pintFailSafeMultiple = 3
		
		pstrSiteURLToRemove = "http://localhost/"
		pstrImagePrefix = "images\"
		pstrImageSuffix = ".jpg"
		pblnUseProdIDForImageName = True
		
		'Display Style
		cstrTableWidth = "100%"
		cstrTableBorder = "0"
		cstrTableCellPadding = "2"
		cstrTableCellSpacing = "0"
		cstrTableBackground = ""
		cstrCellAlign = "center"	'options (left, center, right)
		cstrCellVAlign = "middle"	'options (top, middle, bottom)
		
'//
'////////////////////////////////////////////////////////////////////////////////

pblnIsSQL - set to true if you use a SQL Server database
pblnIsSF5AE - set to true if you use StoreFront 5 AE

pblnDisplayOnlyProductsWithImages - set to true if you want the script to verify
images are present before displaying the product. This will avoid broken images
from appearing in the results

pintFailSafeMultiple - this is used to maximize the returned results if you 
require the images to be present, this multiple is only used if there are a lot
of images referenced that are not present.

pstrSiteURLToRemove - if you have set your images to be absolute paths instead of
relative paths, this should be the path to your web.

pstrImagePrefix - this is the path from your base directory to your image directory
plus any prefix you use before the product ID. It is only used if you wish to use a
default naming scheme for your images. 

pstrImageSuffix  - this is the suffix you use after the product ID to include the
image type extension. It is only used if you wish to use a default naming scheme 
for your images. 

  Ex. if all images are in the smImage directory and named sm_ProdID.jpg then
		pstrImagePrefix = "smImage\sm_"
		pstrImageSuffix = ".jpg"

pblnUseProdIDForImageName - if you use a default naming scheme for you "small" images,
then set this to true. It will only use the the default name if no image path is
set for the small image for the product	

Display parameters - 
set these as desired to configure the table style


2) Create a back-up of any page you wish to display random
   products on. You can display random products on any .asp 
   page on your site by

a) Inserting the line

<!--#include file="SFLib/ssRandomProducts.asp"-->

b) Inserting the line

<% Call WriteRandomProduct %>

where ever you want the random products to display. This line will generate the 
products arranged to your specification enclosed in a table.

OR

<% Call WriteRandomProductBasedOnSearchParamters %>

where ever you want the random products to display filtered based on the querystring. 
This would be used on a customized version of search_results. This line will generate 
the products arranged to your specification enclosed in a table.

Advanced options:

You can limit the products that can be randomly shown by category,
manufacturer, vendor, minimum price, maximum price, keyword, and products
added after a certain date. To do this:

a) Inserting the line

<!--#include file="SFLib/ssRandomProducts.asp"-->

b) Inserting the following code where ever you want the random 
products to display. This will generate the products arranged to 
your specification enclosed in a table. All you need to do is fill
in the respective values.

Note: Manufacturer, Category, and Vendor should be set as their
      respective IDs. If you do not know them, you can view
      advancedsearch.asp and view the source. Find the option
      corresponding to your desired item. It will contain the
      ID number in the value tag
      Ex. <option value="xx">Name</option>
      You can also submit the search and the id will be contained
      in the resulting querystring (URL)
      You may use multiple categories, manufacturers, or vendors. To do so, simply
      separate their IDs by a comma Ex. .CategoryID = "1,3,6"

<% 

Dim pclsRandomProducts

Set pclsRandomProducts = New clsRandomProducts
With pclsRandomProducts
	
	On Error Resume Next
	If isObject(cnn) Then .Connection = cnn
	If Err.number > 0 Then Err.Clear
	On Error Goto 0

	.IsSF5AE = False
	.IsSQL = False
	.NumColumns = 3
	.NumRows = 3
	.DisplayOnlyProductsWithImages = False
	.UseProdIDForImageName = True
	.ManufacturerID = ""
	.CategoryID = ""
	.VendorID = ""
	.SubCategoryID = ""
	.MinPrice = ""
	.MaxPrice = ""
	.AddedAfter = ""
	.DisplayRandomProducts
End With
Set pclsRandomProducts = Nothing

%>



Congratulations! You're done.


