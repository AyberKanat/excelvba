Sub Investor_Count()
'
' denemme Macro
'
Dim ARCount As Integer
'
    Sheets("Investor Count").Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Clear
    Sheets("03-HSD").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("Investor Count").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Investor Count").Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    ARCount = Selection.Rows.Count
    
    'columnfill_B
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Bayi Unvanı"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=+INDEX('01-Bayi Bilgileri'!C10:C28,MATCH('Investor Count'!RC[-1],'01-Bayi Bilgileri'!C[11],0),8)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & ARCount)

    'columnfill_C
    Range("C1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Bölgesi"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=+INDEX('01-Bayi Bilgileri'!C10:C28,MATCH('Investor Count'!RC[-2],'01-Bayi Bilgileri'!C[10],0),10)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & ARCount)

    'columnfill_D
    Range("D1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "BSY"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=+INDEX('01-Bayi Bilgileri'!C10:C28,MATCH('Investor Count'!RC[-3],'01-Bayi Bilgileri'!C[9],0),18)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & ARCount)
   
    'columnfill_E
    Range("E1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "BSD"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=+INDEX('01-Bayi Bilgileri'!C10:C28,MATCH('Investor Count'!RC[-4],'01-Bayi Bilgileri'!C[8],0),19)"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & ARCount)

   
    'columnfill_F
    Range("F1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "PrimTipi"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=+INDEX('01-Bayi Bilgileri'!C10:C28,MATCH('Investor Count'!RC[-5],'01-Bayi Bilgileri'!C[7],0),1)"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F" & ARCount)

   
   
   
    Range("G1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Investor"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(LEFT(RC[-6],7),'02-Yatırımcı Bilgileri'!C[-6]:C[10],13,0)"
    Sheets("02-Yatırımcı Bilgileri").Select
    ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5").Sort. _
        SortFields.Add Key:=Range("Table5[[#All],[Firma Sahibi]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.ListObjects("Table5").Range.AutoFilter Field:=11, Criteria1:= _
        "Hayir"
    Rows("2669:2669").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ListObjects("Table5").Range.AutoFilter Field:=11
    ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5").Sort. _
        SortFields.Add Key:=Range("Table5[[#All],[Yatirimci Kodu Durumu]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.ListObjects("Table5").Range.AutoFilter Field:=14, Criteria1:= _
        "Pasif"
    Rows("2450:2450").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("02-Yatırımcı Bilgileri").ListObjects("Table5").Sort. _
        SortFields.Clear
    Sheets("Investor Count").Select
    Selection.AutoFill Destination:=Range("G2:G" & ARCount)
      Range("H1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Count"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-1],RC[-1],C[-2],RC[-2])"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & ARCount)
End Sub

