Sub CopyAll()
Dim Wb1 As Workbook
Dim Wb2 As Workbook
With Application
.ScreenUpdating = False
.EnableEvents = False
.DisplayAlerts = False
End With
Set Wb1 = Workbooks.Open("c:\temp\Test.xls")
Set Wb2 = ThisWorkbook
Wb1.Sheets.Copy Before:=Wb2.Sheets("00")
Wb1.Close False
With Application
.ScreenUpdating = True
.EnableEvents = True
.DisplayAlerts = True
End With
End Sub