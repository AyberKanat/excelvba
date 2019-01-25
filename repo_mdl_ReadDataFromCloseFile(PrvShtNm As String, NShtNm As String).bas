Sub ReadDataFromCloseFile(PrvShtNm As String, NShtNm As String)
    'On Error GoTo ErrHandler
     Application.DisplayAlerts = False
    'donot forget to restore the standard behavior at the end of your process:

   ' Application.ScreenUpdating = False
    Dim wb1 As Workbook, src As Workbook

    Set wb1 = Application.ActiveWorkbook
    
    Dim GetFileName As String, cellrange As String
        
  'Show the open dialog and pass the selected file name to the String variable "GetFileName"

    cellrange = CStr(NShtNm) & "B"

    GetFileName = Application.GetOpenFilename

    'They have cancelled.

    If GetFileName = "False" Then Exit Sub
    
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(GetFileName, True, True)
    
    ' GET THE TOTAL CBU ROWS FROM THE SOURCE WORKBOOK.
   
    ' COPY CBU DATA FROM SOURCE (CLOSE WORKGROUP) TO THE DESTINATION WORKBOOK.
  '  With wb1
  '      If WorksheetExists(ShtNm) = True Then
  '      wb1.Sheets(ShtNm).Delete
  '      End If
  '  End With
    
    If PrvShtNm = "xx" Then
        src.Activate
        src.Sheets.Copy Before:=wb1.Sheets("StartUp")
    ElseIf PrvShtNm <> "xx" Then
        src.Activate
        src.Sheets(PrvShtNm).Select
        src.Sheets(PrvShtNm).Copy Before:=wb1.Sheets("StartUp")
        wb1.Sheets(PrvShtNm).Select
        wb1.Sheets(PrvShtNm).Name = NShtNm
    End If

    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
 
 Application.DisplayAlerts = True
   
'ErrHandler:
   ' Application.EnableEvents = True
  '  Application.ScreenUpdating = True

End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function
