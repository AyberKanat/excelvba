Sub Adv_Filter()
'
' Adv_Filter Macro
'
     Sheets("XXELIGIBILITYSHEETXX").Select
    Range("A1:BJ5000").Clear
    Range("BYP!A1:BJ20000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "Scheme!BHMCriteria"), CopyToRange:=Range("A1:BJ1"), Unique:=False
End Sub