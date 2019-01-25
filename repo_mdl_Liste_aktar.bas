Sub Liste_aktar()
'
' Liste_aktar Macro
'
    Range("'FSK WOW'!A2:A5000").Clear
    Range([INDIRECT("Eligibility!j5")]).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "Eligibility!j1"), CopyToRange:=Range("'FSK WOW'!A2:A500"), Unique:=False
End Sub
