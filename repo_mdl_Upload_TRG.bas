Sub Upload_TRG()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("TRG") = True Then
        wb1.Sheets("TRG").Delete
    End If
        
        
        'On Error GoTo ErrHandler
     Application.DisplayAlerts = False
    'donot forget to restore the standard behavior at the end of your process:

   ' Application.ScreenUpdating = False
    Dim src As Workbook
    Dim GetFileName As String
    Dim MySheet As String
    
  'Show the open dialog and pass the selected file name to the String variable "GetFileName"

    GetFileName = Application.GetOpenFilename
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(GetFileName, True, True)
    MySheet = src.ActiveSheet.Name
    
    ThisProject.Sheets("Scheme").Select
    Range("CC1").Select

    ActiveCell.Offset(0, 0).Value = MySheet
    
    ' GET THE TOTAL CBU ROWS FROM THE SOURCE WORKBOOK.

    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
 
 Application.DisplayAlerts = True

    ThisProject.ActiveWorkbook.Sheets("Scheme").Select
    Range("CC1").Activate
    MySheet = ActiveCell.Offset(0, 0).Value
    
  Call ReadDataFromCloseFile(MySheet, "TRG")
     Sheets("StartUp").Select
    Range("A1").Select
End Sub

