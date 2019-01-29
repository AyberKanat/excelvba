Sub Macro1()
'
' Macro1 Macro
'
Dim RCount As Integer
'
    Sheets("Calculation").Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Clear
    Sheets("Investor HG").Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    RCount = Selection.Rows.Count
    Selection.Copy
    Sheets("Calculation").Select
    Range("A1").Select
    ActiveSheet.Paste
End Sub
