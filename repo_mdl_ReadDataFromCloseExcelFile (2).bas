Sub ReadDataFromCloseFile(PrvShtNm As String, NShtNm As String)
    'On Error GoTo ErrHandler
     Application.DisplayAlerts = False
    'do not forget to restore the standard behavior at the end of your process:

   ' Application.ScreenUpdating = False
    Dim wb1 As Workbook, src As Workbook


    Set wb1 = Application.ActiveWorkbook
    
    Dim GetFileName As String
        
  'Show the open dialog and pass the selected file name to the String variable "GetFileName"


    GetFileName = Application.GetOpenFilename

    'They have cancelled.

    If GetFileName = "False" Then Exit Sub
    
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(GetFileName, True, True)
    
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK.
   
    
    src.Activate
    src.Sheets(PrvShtNm).Select
    src.Sheets(PrvShtNm).Copy Before:=wb1.Sheets("XXOPENINGSHEETXX")
    wb1.Sheets(PrvShtNm).Select
    wb1.Sheets(PrvShtNm).Name = NShtNm
    
    
   
 
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
 
 Application.DisplayAlerts = True
   
'ErrHandler:
   ' Application.EnableEvents = True
  '  Application.ScreenUpdating = True


End Sub
