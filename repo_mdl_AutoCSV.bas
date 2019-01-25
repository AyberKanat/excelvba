Option Compare Database 
Option Explicit 

Function DoImport() 

 Dim strPathFile As String 
 Dim strFile As String 
 Dim strPath As String 
 Dim strTable As String 
 Dim blnHasFieldNames As Boolean 

 ' Change this next line to True if the first row in CSV worksheet 
 ' has field names 
 blnHasFieldNames = True 

 ' Replace C:\Documents\ with the real path to the folder that 
 ' contains the CSV files 
 strPath = "C:\Documents\" 

 ' Replace tablename with the real name of the table into which 
 ' the data are to be imported 

 strFile = Dir(strPath & "*.csv") 


 Do While Len(strFile) > 0 
       strTable = Left(strFile, Len(strFile) - 4) 
       strPathFile = strPath & strFile 
       DoCmd.TransferText acImportDelim, , strTable, strPathFile, blnHasFieldNames 


 ' Uncomment out the next code step if you want to delete the 
 ' EXCEL file after it's been imported 
 '       Kill strPathFile 

       strFile = Dir() 

 Loop 


End Function 
