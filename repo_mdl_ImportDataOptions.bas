' Import Data from All EXCEL Files in a single Folder via TransferSpreadsheet (VBA)

' Generic code to import the data from the first (or only) worksheet in all EXCEL files that are located within a single folder. All of the EXCEL files' worksheets must have the data in the same layout and format.

Dim strPathFile As String, strFile As String, strPath As String
Dim strTable As String
Dim blnHasFieldNames As Boolean

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

' Replace C:\Documents\ with the real path to the folder that
' contains the EXCEL files
strPath = "C:\Documents\"

' Replace tablename with the real name of the table into which 
' the data are to be imported
strTable = "tablename"

strFile = Dir(strPath & "*.xls")
Do While Len(strFile) > 0
      strPathFile = strPath & strFile
      DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
            strTable, strPathFile, blnHasFieldNames

' Uncomment out the next code step if you want to delete the 
' EXCEL file after it's been imported
'       Kill strPathFile

      strFile = Dir()
Loop

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Import Data from Specific Worksheets in All EXCEL Files in a single Folder via TransferSpreadsheet (VBA)

' Generic code to import the data from specific worksheets in all EXCEL files (worksheet names are the same in all files) that are located within a single folder. All of the EXCEL files' worksheets with the same worksheet names must have the data in the same layout and format.

Dim strPathFile As String, strFile As String, strPath As String
Dim blnHasFieldNames As Boolean
Dim intWorksheets As Integer

' Replace 3 with the number of worksheets to be imported
' from each EXCEL file
Dim strWorksheets(1 To 3) As String

' Replace 3 with the number of worksheets to be imported
' from each EXCEL file (this code assumes that each worksheet
' with the same name is being imported into a separate table 
' for that specific worksheet name)
Dim strTables(1 To 3) As String

' Replace generic worksheet names with the real worksheet names;
' add / delete code lines so that there is one code line for
' each worksheet that is to be imported from each workbook file
strWorksheets(1) = "GenericWorksheetName1"
strWorksheets(2) = "GenericWorksheetName2"
strWorksheets(3) = "GenericWorksheetName3"

' Replace generic table names with the real table names;
' add / delete code lines so that there is one code line for
' each worksheet that is to be imported from each workbook file
strTables(1) = "GenericTableName1"
strTables(2) = "GenericTableName2"
strTables(3) = "GenericTableName3"

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

' Replace C:\Documents\ with the real path to the folder that
' contains the EXCEL files
strPath = "C:\Documents\"

' Replace 3 with the number of worksheets to be imported
' from each EXCEL file
For intWorksheets = 1 To 3

      strFile = Dir(strPath & "*.xls")
      Do While Len(strFile) > 0
            strPathFile = strPath & strFile
            DoCmd.TransferSpreadsheet acImport, _
                  acSpreadsheetTypeExcel9, strTables(intWorksheets), _
                  strPathFile, blnHasFieldNames, _
                  strWorksheets(intWorksheets) & "$"
            strFile = Dir()
      Loop

Next intWorksheets

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Import Data from All Worksheets in a single EXCEL File into One Table via TransferSpreadsheet (VBA)

' Generic code to import the data from all worksheets in a single EXCEL file. Because all of the worksheets' data will be imported into the same table, all of the EXCEL files' worksheets must have the data in the same layout and format.

Dim blnHasFieldNames As Boolean, blnEXCEL As Boolean, blnReadOnly As Boolean
Dim lngCount As Long
Dim objExcel As Object, objWorkbook As Object
Dim colWorksheets As Collection
Dim strPathFile as String, strTable as String
Dim strPassword As String

' Establish an EXCEL application object
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
      Set objExcel = CreateObject("Excel.Application")
      blnEXCEL = True
End If
Err.Clear
On Error GoTo 0

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

' Replace C:\Filename.xls with the actual path and filename
strPathFile = "C:\Filename.xls"

' Replace tablename with the real name of the table into which 
' the data are to be imported
strTable = "tablename"

' Replace passwordtext with the real password;
' if there is no password, replace it with vbNullString constant
' (e.g., strPassword = vbNullString)
strPassword = "passwordtext"

blnReadOnly = True ' open EXCEL file in read-only mode

' Open the EXCEL file and read the worksheet names into a collection
Set colWorksheets = New Collection
Set objWorkbook = objExcel.Workbooks.Open(strPathFile, , blnReadOnly, , _
      strPassword)
For lngCount = 1 To objWorkbook.Worksheets.Count
      colWorksheets.Add objWorkbook.Worksheets(lngCount).Name
Next lngCount

' Close the EXCEL file without saving the file, and clean up the EXCEL objects
objWorkbook.Close False
Set objWorkbook = Nothing
If blnEXCEL = True Then objExcel.Quit
Set objExcel = Nothing

' Import the data from each worksheet into the table
For lngCount = colWorksheets.Count To 1 Step -1
      DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
            strTable, strPathFile, blnHasFieldNames, colWorksheets(lngCount) & "$"
Next lngCount

' Delete the collection
Set colWorksheets = Nothing

' Uncomment out the next code step if you want to delete the 
' EXCEL file after it's been imported
' Kill strPathFile

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Import Data from All Worksheets in a single EXCEL File into Separate Tables via TransferSpreadsheet (VBA)

' Generic code to import the data from all worksheets in a single EXCEL file. Each worksheet's data will be imported into a separate table whose name is 'tbl' plus the worksheet name (e.g., "tblSheet1").

Dim blnHasFieldNames As Boolean, blnEXCEL As Boolean, blnReadOnly As Boolean
Dim lngCount As Long
Dim objExcel As Object, objWorkbook As Object
Dim colWorksheets As Collection
Dim strPathFile As String
Dim strPassword As String

' Establish an EXCEL application object
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
      Set objExcel = CreateObject("Excel.Application")
      blnEXCEL = True
End If
Err.Clear
On Error GoTo 0

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

' Replace C:\Filename.xls with the actual path and filename
strPathFile = "C:\Filename.xls"

' Replace passwordtext with the real password;
' if there is no password, replace it with vbNullString constant
' (e.g., strPassword = vbNullString)
strPassword = "passwordtext"

blnReadOnly = True ' open EXCEL file in read-only mode

' Open the EXCEL file and read the worksheet names into a collection
Set colWorksheets = New Collection
Set objWorkbook = objExcel.Workbooks.Open(strPathFile, , blnReadOnly, , _
      strPassword)
For lngCount = 1 To objWorkbook.Worksheets.Count
      colWorksheets.Add objWorkbook.Worksheets(lngCount).Name
Next lngCount

' Close the EXCEL file without saving the file, and clean up the EXCEL objects
objWorkbook.Close False
Set objWorkbook = Nothing
If blnEXCEL = True Then objExcel.Quit
Set objExcel = Nothing

' Import the data from each worksheet into a separate table
For lngCount = colWorksheets.Count To 1 Step -1
      DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
            "tbl" & colWorksheets(lngCount), strPathFile, blnHasFieldNames, _
            colWorksheets(lngCount) & "$"
Next lngCount

' Delete the collection
Set colWorksheets = Nothing

' Uncomment out the next code step if you want to delete the 
' EXCEL file after it's been imported
' Kill strPathFile

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Import Data from A Specific Worksheet in All EXCEL Files in a single Folder into Separate Tables via TransferSpreadsheet (VBA)

' Generic code to import the data from a specific worksheet in all EXCEL files in a single folder. Each worksheet's data will be imported into a separate table whose name is 'tbl_' plus the workbook name without the ".xls" file extension (e.g., "tbl_NameOfFile").

Dim blnHasFieldNames as Boolean
Dim strWorksheet As String, strTable As String
Dim strPath As String, strPathFile As String

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

' Replace C:\Documents\ with the real path to the folder that
' contains the EXCEL files
strPath = "C:\Documents\"

' Replace worksheetname with the real name of the worksheet that is to be
' imported from each file
strWorksheet = "worksheetname"

' Import the data from each workbook file in the folder
strFile = Dir(strPath & "*.xls")
Do While Len(strFile) > 0
      strPathFile = strPath & strFile
      strTable = "tbl_" & Left(strFile, InStrRev(strFile, ".xls") - 1)

      DoCmd.TransferSpreadsheet acImport, _
            acSpreadsheetTypeExcel9, strTable, strPathFile, _
            blnHasFieldNames, strWorksheet & "$"

      ' Uncomment out the next code step if you want to delete the 
      ' EXCEL file after it's been imported
      ' Kill strPathFile

      strFile = Dir()
Loop

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Import Data from All Worksheets in All EXCEL Files in a single Folder into Separate Tables via TransferSpreadsheet (VBA)

' Generic code to import the data from all worksheets in all EXCEL files in a single folder. Each worksheet's data will be imported into a separate table whose name is 'tbl' plus the worksheet name plus an integer value that represents a "counter" for the workbooks (e.g., "tblWorksheetName1").

Dim blnHasFieldNames As Boolean, blnEXCEL As Boolean, blnReadOnly As Boolean
Dim intWorkbookCounter As Integer
Dim lngCount As Long
Dim objExcel As Object, objWorkbook As Object
Dim colWorksheets As Collection
Dim strPath As String, strFile As String
Dim strPassword As String

' Establish an EXCEL application object
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
      Set objExcel = CreateObject("Excel.Application")
      blnEXCEL = True
End If
Err.Clear
On Error GoTo 0

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

' Replace C:\MyFolder\ with the actual path to the folder that holds the EXCEL files
strPath = "C:\MyFolder\"

' Replace passwordtext with the real password;
' if there is no password, replace it with vbNullString constant
' (e.g., strPassword = vbNullString)
strPassword = "passwordtext"

blnReadOnly = True ' open EXCEL file in read-only mode

strFile = Dir(strPath & "*.xls")

intWorkbookCounter = 0

Do While strFile <> ""

      intWorkbookCounter = intWorkbookCounter + 1

      Set colWorksheets = New Collection

      Set objWorkbook = objExcel.Workbooks.Open(strPath & strFile, , _
            blnReadOnly, , strPassword)

      For lngCount = 1 To objWorkbook.Worksheets.Count
            colWorksheets.Add objWorkbook.Worksheets(lngCount).Name
      Next lngCount

      ' Close the EXCEL file without saving the file, and clean up the EXCEL objects
      objWorkbook.Close False
      Set objWorkbook = Nothing

      ' Import the data from each worksheet into a separate table
      For lngCount = colWorksheets.Count To 1 Step -1
            DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
                  "tbl" & colWorksheets(lngCount) & intWorkbookCounter, _
                  strPath & strFile, blnHasFieldNames, _
                  colWorksheets(lngCount) & "$"
      Next lngCount

      ' Delete the collection
      Set colWorksheets = Nothing

      ' Uncomment out the next code step if you want to delete the 
      ' EXCEL file after it's been imported
      ' Kill strPath & strFile

      strFile = Dir()

Loop

If blnEXCEL = True Then objExcel.Quit
Set objExcel = Nothing

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Browse to a single EXCEL File and Import Data from that EXCEL File via TransferSpreadsheet (VBA)

' Generic code to browse to a single EXCEL file, and then to import the data from the first (or only) worksheet in that EXCEL file. This generic method uses the Windows API to browse to a single file; the code for this API (which was written by Ken Getz) is located at The ACCESS Web ( http://theaccessweb.com/ ).

' First step is to paste all the Getz code (from http://theaccessweb.com/api/api0001.htm ) into a new, regular module in your database. Be sure to give the module a unique name (i.e., it cannot have the same name as any other module, any other function, or any other subroutine in the database). Then use this generic code to allow the user to select the EXCEL file that is to be imported.

Dim strPathFile As String
Dim strTable As String, strBrowseMsg As String
Dim strFilter As String, strInitialDirectory As String
Dim blnHasFieldNames As Boolean

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

strBrowseMsg = "Select the EXCEL file:"

' Change C:\MyFolder\ to the path for the folder where the Browse
' window is to start (the initial directory). If you want to start in
' ACCESS' default folder, delete C:\MyFolder\ from the code line,
' leaving an empty string as the value being set as the initial
' directory
strInitialDirectory = "C:\MyFolder\"

strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.xls)", "*.xls")

strPathFile = ahtCommonFileOpenSave(InitialDir:=strInitialDirectory, _
      Filter:=strFilter, OpenFile:=False, _
      DialogTitle:=strBrowseMsg, _
      Flags:=ahtOFN_HIDEREADONLY)

If strPathFile = "" Then
      MsgBox "No file was selected.", vbOK, "No Selection"
      Exit Sub
End If

' Replace tablename with the real name of the table into which 
' the data are to be imported
strTable = "tablename"

DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
      strTable, strPathFile, blnHasFieldNames

' Uncomment out the next code step if you want to delete the 
' EXCEL file after it's been imported
' Kill strPathFile

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Browse to a single Folder and Import Data from All EXCEL Files in that Folder via TransferSpreadsheet (VBA)

' Generic code to browse to a single folder, and then to import the data from the first (or only) worksheet in all EXCEL files that are located within that folder. All of the EXCEL files' worksheets must have the data in the same layout and format. This generic method uses the Windows API to browse to a single folder; the code for this API (which was written by Terry Kreft) is located at The ACCESS Web ( http://theaccessweb.com/ ).

' First step is to paste all the Kreft code (from http://theaccessweb.com/api/api0002.htm ) into a new, regular module in your database. Be sure to give the module a unique name (i.e., it cannot have the same name as any other module, any other function, or any other subroutine in the database). Then use this generic code to allow the user to select the folder in which the EXCEL files are located.

Dim strPathFile As String, strFile As String, strPath As String
Dim strTable As String, strBrowseMsg As String
Dim blnHasFieldNames as Boolean

' Change this next line to True if the first row in EXCEL worksheet
' has field names
blnHasFieldNames = False

strBrowseMsg = "Select the folder that contains the EXCEL files:"

strPath = BrowseFolder(strBrowseMsg)

If strPath = "" Then
      MsgBox "No folder was selected.", vbOK, "No Selection"
      Exit Sub
End If

' Replace tablename with the real name of the table into which 
' the data are to be imported
strTable = "tablename"

strFile = Dir(strPath & "\*.xls")
Do While Len(strFile) > 0
      strPathFile = strPath & "\" & strFile
      DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, _
            strTable, strPathFile, blnHasFieldNames

' Uncomment out the next code step if you want to delete the 
' EXCEL file after it's been imported
'       Kill strPathFile

      strFile = Dir()
Loop

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Read Data from EXCEL File via Query (SQL Statement)

' Generic SQL statement that reads data from an EXCEL file. Replace C:\MyFolder\MyFile.xls with the real path and filename of the EXCEL file. Replace WorksheetName with the real name of the worksheet -- NOTE that the name cannot be longer than 30 characters (one less than EXCEL's limit for a worksheet name) or else ACCESS / Jet will give you an error stating that the file cannot be found. In this SQL statement, HDR=YES means that the first row of data are header names (change to NO if the first row does not contain header names); IMEX=1 alllows "mixed formatting" within a column (alpha characters and numbers, for example) so that errors will not be raised when importing mixed formats; the $ character must be immediately after the worksheet name; and A2:U66536 is the range of data to be imported (these cell references can be changed to any contiguous range of cells in the worksheet).

SELECT T1.*, 1 AS SheetSource
FROM [Excel 8.0;HDR=YES;IMEX=1;Database=C:\MyFolder\MyFile.xls].[WorksheetName$A2:U65536] as T1;

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Write Data From an EXCEL Worksheet into a Recordset using Automation (VBA)

' Generic code to open a recordset (based on an existing table) for the data that are to be imported from a worksheet in an EXCEL file, and then to loop through the recordset and write each cell's value into a field in the recordset, with each row in the worksheet being written into a separate record. The starting cell for the EXCEL worksheet is specified in the code; after that, the data are read from contiguous cells and rows. This code example uses "late binding" for the EXCEL automation, and this code assumes that the EXCEL worksheet DOES NOT contain header information in the first row of data being read.

Dim lngColumn As Long
Dim xlx As Object, xlw As Object, xls As Object, xlc As Object
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim blnEXCEL As Boolean

blnEXCEL = False

' Establish an EXCEL application object
On Error Resume Next
Set xlx = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
      Set xlx = CreateObject("Excel.Application")
      blnEXCEL = True
End If
Err.Clear
On Error GoTo 0

' Change True to False if you do not want the workbook to be
' visible when the code is running
xlx.Visible = True

' Replace C:\Filename.xls with the actual path and filename
' of the EXCEL file from which you will read the data
Set xlw = xlx.Workbooks.Open("C:\Filename.xls", , True) ' opens in read-only mode

' Replace WorksheetName with the actual name of the worksheet
' in the EXCEL file
Set xls = xlw.Worksheets("WorksheetName")

' Replace A1 with the cell reference from which the first data value
' (non-header information) is to be read
Set xlc = xls.Range("A1") ' this is the first cell that contains data

Set dbs = CurrentDb()

' Replace QueryOrTableName with the real name of the table or query
' that is to receive the data from the worksheet
Set rst = dbs.OpenRecordset("QueryOrTableName", dbOpenDynaset, dbAppendOnly)

' write data to the recordset
Do While xlc.Value <> ""
      rst.AddNew
            For lngColumn = 0 To rst.Fields.Count - 1
                  rst.Fields(lngColumn).Value = xlc.Offset(0, lngColumn).Value
            Next lngColumn
      rst.Update
      Set xlc = xlc.Offset(1,0)
Loop

rst.Close
Set rst = Nothing

dbs.Close
Set dbs = Nothing

' Close the EXCEL file without saving the file, and clean up the EXCEL objects
Set xlc = Nothing
Set xls = Nothing
xlw.Close False
Set xlw = Nothing
If blnEXCEL = True Then xlx.Quit
Set xlx = Nothing

' Return to Top of Page

' Return to EXCEL Main Page

' Return to Home

 

' Avoid DataType Mismatch Errors when Importing Data from an EXCEL File or when Linking to an EXCEL File

' When importing data from an EXCEL spreadsheet into an ACCESS table via the TransferSpreadsheet action, or when linking to an EXCEL spreadsheet as a linked ACCESS table, often you will see the "#Num!" error code for the value in a field in the ACCESS table; or you will see that leading zeroes are lost from text strings that contain only number characters; or you will see that text strings longer than 255 characters are truncated in a field in the ACCESS table.

' The "#Num!" error code that you see is because Jet (ACCESS) sees only numeric values in the first 8 - 25 rows of data in the EXCEL sheet, even though you have formatted the EXCEL column as "Text". In EXCEL, if you change the format from "General" or a numeric format to "Text", the previous numeric format for a cell will "stick" to numeric values.

' What ACCESS and Jet are doing is assuming that the "text" data actually are numeric data, and thus all your non-numeric text strings are "not matching" to a numeric data type. One of these suggestions should fix the problem:

' 1) Insert a ' (apostrophe) character at the beginning of each cell's value for that column in the EXCEL file -- that should let Jet (ACCESS) treat that column's values as text and not numeric.

' 2) Insert a dummy row of data as the first row, where the dummy row contains nonnumeric characters in the cell in that column -- that should let Jet (ACCESS) treat that column's values as text and not numeric.

' 3) Double-click into the EXCEL cell that has the "numeric" data, then click on any other cell -- that will "update" the cell to the "Text" format.

' 4) Create a blank table into which you will import the spreadsheet's data. For the field that will receive the numeric data, make its data type "Text". Jet (ACCESS) then will "honor" the field's datatype when it does the import.

' The loss of leading zeroes from text strings that contain only number characters is a symptom of the same problem noted above for the "#Num!" error code. One of the these suggestions should fix the problem:

' 1) Insert a ' (apostrophe) character at the beginning of each cell's value for that column in the EXCEL file -- that should let Jet (ACCESS) treat that column's values as text and not numeric.

' 2) Insert a dummy row of data as the first row, where the dummy row contains nonnumeric characters in the cell in that column -- that should let Jet (ACCESS) treat that column's values as text and not numeric.

' 3) Create a blank table into which you will import the spreadsheet's data. For the field that will receive the numeric data, make its data type "Text". Jet (ACCESS) then will "honor" the field's datatype when it does the import.

' The truncated text string that you see is because Jet (ACCESS) sees only "short text" (text strings no longer than 255 characters) values in the first 8 - 25 rows of data in the EXCEL sheet, even though you have longer text farther down the rows. What ACCESS and Jet are doing is assuming that the "text" data actually are Text data type, not Memo data type. One of these suggestions should fix the problem:

' 1) Insert a dummy row of data as the first row, where the dummy row contains a text string longer than 255 characters in the cell in that column -- that should let Jet (ACCESS) treat that column's values as memo and not text.

' 2) Create a blank table into which you will import the spreadsheet's data. For the field that will receive the "memo" data, make its data type "Memo". Jet (ACCESS) then will "honor" the field's datatype when it does the import.

' It's possible to force Jet to scan all the rows and not guess the data type based on just the first few rows. See this article for information about the registry key (see TypeGuessRows and MaxScanRows information):  http://dailydoseofexcel.com/archives/2004/06/03/external-data-mixed-data-types/    [ NOTE: There are some reports by others that this registry key may not work as expected when using Windows XP SP3 or when using ACCESS 2007. ]
