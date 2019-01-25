Sub ListWorkSheetNamesNewWs()
'Updateby20140624
Dim xWs As Worksheet
On Error Resume Next
Application.DisplayAlerts = False
xTitleId = "XXOPENINGSHEETXX"
'Application.Sheets(xTitleId).Delete
'Application.Sheets.Add Application.Sheets(1)
Set xWs = Application.Sheet("XXOPENINGSHEETXX")
xWs.Name = xTitleId
For i = 2 To Application.Sheets.Count
    xWs.Range("XXLISTCOLUMNXX" & (i - 1)) = Application.Sheets(i).Name
Next
Application.DisplayAlerts = True
End Sub
