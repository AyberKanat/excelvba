Sub Upload_XXYOURSHEETXX()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("XXYOURSHEETXX") = True Then
        wb1.Sheets("XXYOURSHEETXX").Delete
    End If
  ReadDataFromCloseFile("XXSOURCESHEETXX", "XXYOURSHEETXX")
     Sheets("XXOPENINGSHEETXX").Select
    Range("A1").Select
End Sub
