Sub Macro2()
'
' Macro2 Macro
'

'
    Sheets("Scheme").Select
    Range("I1:I45").Select
    ActiveWorkbook.Worksheets("Scheme").Names.Add Name:="dennnen", RefersToR1C1 _
        :="=Scheme!R1C9:R45C9"
    ActiveWorkbook.Worksheets("Scheme").Names("dennnen").Comment = ""
End Sub
