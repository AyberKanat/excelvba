Sub Macro2()
'
' Macro2 Macro
'

'
    Columns("E:H").Select
    Cells.Replace What:="#REF!", Replacement:="'03-HSD!'", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="#REF!", Replacement:="'03-HSD'", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="#REF!", Replacement:="03-HSD", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="#REF", Replacement:="03-HSD", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("F6").Select
End Sub
Sub Macro3()
'
' Macro3 Macro
'

'
    Columns("E:H").Select
    Selection.Replace What:="#REF", Replacement:="'03-HSD'", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
