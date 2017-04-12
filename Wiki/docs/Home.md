**Project Description**
This project is to contain a number of small macros that provide often asked for functionality, written for various versions of MS Office.

Programming Language: VBA (Visual Basic for Applications)
Language: English (en-GB)
Related Software: Microsoft Office, Excel, Word
# Macros
* **[Worksheet Merge Macro](#WorksheetMergeMacro)**
* **[Common Excel Functions](#CommonExcelFunctions)**
* **[Google Translate / Web Service Caller](#GoogleTranslate)**
* **[Word SaveAs with Auto Naming Convention](#SaveAsAutoName)**

# Donations
If you'd like to show your appreciation for what's on here, please take a look at this thread ([http://officemacros.codeplex.com/discussions/251401](http://officemacros.codeplex.com/discussions/251401)) on the discussions page. Though I don't accept monetary donations personally, if you'd like to say thanks by giving to a worthy cause, that would be much appreciated by me.  If you'd like to contribute code to the project, or have something suitable which you'd like me to link to, please get in touch.

{anchor:WorksheetMergeMacro}
# Worksheet Merge Macro
## The Problem
You have two lists which share a key.  You want to match the two up so that the row with a given key from sheet 1 will be against the same row number as the row from sheet 2 with the same key.  The lists may contain items which do not have a corresponding item in the other list, or there may be a different number of items under each key for each list.
## The Solution
The macro makes a copy of both lists, then sorts them in order of the shared key.  Next, the macro goes through each of the lists, leaving rows in place where list1's key matches list2's key for row N, or inserting a blank row into one of the lists where that list does not have an item which appears in the other list.
## How to Use
# Check that your excel security settings allow macros (tools, macros, security - select either medium or low)
# Open the workbook(s) on which your lists appear.
# Open WorkSheetMergeMacro.xls
# Enable macros if prompted
# On the Sheet1 Tab of the Merge form, select the workbook on which one of your lists appears.
# Select the worksheet on which this list appears
# If Row1 contains column headers, tick the Header Row box; otherwise leave it blank
# Select the sort column / the column which contains the shared key data.
# Go to the Sheet2 Tab
# Select the workbook, worksheet, and sort (key) column for your other list
# Click Merge
# A new workbook should have now been created containing sorted copies of the two lists.  Copying and pasting one list next to the other will show that these are now matched up using the shared key.
## Issues
I wrote this macro many years ago, and am aware that there are a few small issues with it.  I'll fix these if this download proves popular.
* The LAST{"_"}ROW and LAST{"_"}COLUMN values are hard coded, rather than automatically worked out.  
* Header Row only works for the first row on a sheet
* The merged sheets are left on separate sheets, rather than providing an option to automatically merge the data onto one sheet.
{anchor:CommonExcelFunctions}
# Common Excel Functions
## Summary
Various functions useful when developing macros:
* =MaxRows()
	* Returns the total number of rows on the spreadsheet (65536 for Excel 97/2003 spreadsheets, 1048576 for Excel 2007 and above).
* =MaxCols()
	* Returns the total number of columns on the spreadsheet (256 for Excel 97/2003 spreadsheets, 16384 for Excel 2007 and above).
* =GetCellAddress(3,7)
	* Returns a fixed cell address given the cell's coordinates (e.g. row 3, column 7 returns $G$3).
* =GetColumnName(5)
	* Returns the alpha code for the given column number (e.g. column 5 returns E).
* =LastPopulatedRow()
	* Returns the last row on the sheet containing data, regardless of which column it appears in.
* =LastRowInRange(E5:K17)
	* Returns the row number of the last row in the given range (so in the example range E5:K17 that would be row 17).
* =LastPopulatedRowInRange(E5:K17)
	* As with LastPopulatedRow(), except searching within the given range rather than the whole sheet (so in the example range E5:K17, the return value could only be a value between 5 and 17, or #NAME? if no cell in the range contains data).
* =LastPopulatedCol()
	* Returns the last col on the sheet containing data, regardless of which row it appears in.
* =LastColInRange(E5:K17)
	* Returns the column number of the last column in the given range (so in the example range E5:K17 that would be column 11 (K)).
* =LastPopulatedColInRange(E5:K17)
	* As with LastPopulatedCol(), except searching within the given range rather than the whole sheet (so in the example range E5:K17, the return value could only be a value between 5 (E) and 11 (K), or #NAME? if no cell in the range contains data).
* =FirstPopulatedRow()
	* As with LastPopulatedRow, but the first one.
* =FirstRowInRange(E5:K17)
	* As with LastRowInRange, but the first one (e.g. 5).
* =FirstPopulatedRowInRange(E5:K17)
	* As with LastPopulatedRowInRange, but the first one.
* =FirstPopulatedCol()
	* As with LastPopulatedCol, but the first one.
* =FirstColInRange(E5:K17)
	* As with LastColInRange, but the first one (e.g. 5 (E)).
* =FirstPopulatedColInRange(E5:K17)
	* As with LastPopulatedCol, but the first one.
{anchor:GoogleTranslate}
# Google Translate / Web Service Caller
## Summary
This allows you to call web services from formulas in excel. 
A function to call the Google Translate service has been included, which wraps the logic needed to populate the parameters, and extract the translation from the returned JSON.  
* =Translate(A1,"en","fr")
	* Translates the contents of cell A1 from English to French.
* =CallWebService("http://www.google.co.uk")
	* returns the html for Google's UK site.
{anchor:SaveAsAutoName}
# Word SaveAs with Auto Naming Convention
## Summary
A word template which overrides the native Save and SaveAs commands to use a custom filename format.
The filename format is the first line of text from the document, followed by the date (in YYYYMMDD format).
If the first line of text is more than 32 chars, the first sentence is used (i.e. up to the first full-stop); again suffixed by the date.
If the first sentence is over 32 chars, the first 32 chars are used; again suffixed by the date.
Hopefully you should find it easy to adapt this macro to your naming needs, but any problems, just let me know & I'll try to help out.
