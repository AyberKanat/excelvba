Sub Adv_Filter()
'
' Adv_Filter Macro
'
    Sheets("Investor HG").Select
    Range("A2:H5000").Clear
    Range("'Investor Count'!A1:H5000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("Main!Z1:AA4"), CopyToRange:=Range("A1:H1"), Unique:=False
    Sheets("Investor HG").Select
    Columns("I:L").Select
    Selection.Replace What:="#REF", Replacement:="'03-HSD'", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub
