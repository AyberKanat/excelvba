Private Sub Workbook_Open()
    Call ReadDataFromCloseFile
End Sub

Sub ReadDataFromCloseFile()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open("C:\Q-SALES.xlsx", True, True)
    
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK.
    Dim iTotalRows As Integer
    iTotalRows = src.Worksheets("sheet1").Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
   
    ' COPY DATA FROM SOURCE (CLOSE WORKGROUP) TO THE DESTINATION WORKBOOK.
    Dim iCnt As Integer         ' COUNTER.
    For iCnt = 1 To iTotalRows
        Worksheets("Sheet1").Range("B" & iCnt).Formula =   
            src.Worksheets("Sheet1").Range("B" & iCnt).Formula
    Next iCnt
    
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub