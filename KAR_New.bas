Private Function NewWorkbook(excl_name As String) As Workbook
Dim ws As Worksheet
Dim fcn As Variant, frn As Variant
Dim lcol As Variant, lrow As Variant, lcoln As String
    Application.DisplayAlerts = False
    Set NewWorkbook = Workbooks.Add
    With NewWorkbook
        .Title = "BHM Mar 18 " + " (" + Format(Date, "mmmm_yyyy") + ")"
        .Subject = "BHM Mar 18 " + " (" + Format(Date, "mmmm_yyyy") + ")"
        .SaveAs Filename:="C:\Users\akanat\Desktop\AK\wowyayın" + "\" + Format(Date, "dd mmmm yyyy") + "_" + "BHM Mart  18 BHM Kontrol v1" + "_" + excl_name + ".xlsx"
    End With
    ThisWorkbook.Activate
    For Each ws In Worksheets
        If ws.Tab.ColorIndex = 1 Or ws.Tab.ColorIndex = 3 Then
            ws.Copy Before:=NewWorkbook.Sheets("Sheet1")
        End If
    Next ws
    Sheets("Sheet1").Delete
    For Each ws In Worksheets
        ws.Activate
        If ws.Tab.ColorIndex = 3 Then
            ws.AutoFilterMode = False
            fcn = Find_Column("Report")
            frn = Find_Row("Report")
            lcol = Cells(frn, "IV").End(xlToLeft).Column
            lcoln = Split(Cells(frn, lcol).Address, "$")(1)
            lrow = Cells(frn, "A").End(xlDown).Row
            Selection.AutoFilter
            Selection.AutoFilter Field:=fcn
            Selection.AutoFilter Field:=fcn, Criteria1:="<>" & excl_name & ""
            ws.Range("A" & frn + 1, lcoln & lrow).Select
        End If
        Selection.Delete
        ws.AutoFilterMode = False
    Next ws
    NewWorkbook.Save
    NewWorkbook.Close False
    Set NewWorkbook = Nothing
    Application.ScreenUpdating = True
End Function

Private Function Find_Column(findstr As String)
    Dim FindString As String
    Dim Rng As Range
    FindString = findstr
    If Trim(FindString) <> "" Then
        With ActiveSheet.UsedRange
            Set Rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                Application.GoTo Rng, True
            Else
                MsgBox "Nothing found"
            End If
        End With
    End If
Find_Column = Rng.Column
End Function

Private Function Find_Row(findstr As String)
    Dim FindString As String
    Dim Rng As Range
    FindString = findstr
    If Trim(FindString) <> "" Then
        With ActiveSheet.UsedRange
            Set Rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                Application.GoTo Rng, True
            Else
                MsgBox "Nothing found"
            End If
        End With
    End If
Find_Row = Rng.Row
End Function


Public Sub KAR()
Dim lcol As Variant, lrow As Variant
ThisWorkbook.Sheets("Main").Activate
lrow = Cells(1, "A").End(xlDown).Row
For i = 2 To lrow
    Range("A" & i).Select
    Call NewWorkbook(Selection.Value)
Next i
End Sub
