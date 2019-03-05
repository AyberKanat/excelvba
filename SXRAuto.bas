Public Function KPI_HG(ByVal Actual As Double, ByVal Target As Double) As Double

    KPI_HG = IIf(Target < 1, 0, (Actual / Target))

End Function

Public Function KPI_Points(ByVal KPI_HGcl As Double, ByVal Limit1 As Double, ByVal Limit2 As Double, ByVal Limit3 As Double, ByVal Points1 As Double, ByVal Points2 As Double, ByVal Points3 As Double) As Double

If KPI_HGcl >= Limit1 Then KPI_Points = Points1
    If KPI_HGcl < Limit1 And KPI_HGcl >= Limit2 Then KPI_Points = Points2
        If KPI_HGcl < Limit2 And KPI_HGcl >= Limit3 Then KPI_Points = Points3
            If KPI_HGcl < Limit3 Then KPI_Points = 0
End Function

Public Function KPI_LmtvsPnt(ByVal KPI_HGcl As Double, ByVal NoOfLimits As Double, limit As Variant, Points As Variant) As Double
' KPI_HGcl lookup Target Realization value for KPI, NoOfLimits how many points brackets are present excluding no points, limit is the array of minimum target realization limits, Points is the array of corresponding Points
Dim i As Double
For i = 1 To NoOfLimits
    If KPI_HGcl >= limit(i) Then Exit For
    If KPI_HGcl < limit(i) Then
    End If
Next i
KPI_LmtvsPnt = Points(i)


Public Sub apply_Error_Control()
    Dim cel As Range

    For Each cel In Selection
        If cel.HasFormula Then
            'option 1
            cel.Formula = Replace(cel.Formula, "=", "=IFERROR(", 1, 1) & ",0)"
            'option 2
            'cel.Formula = "=IFERROR(" & Mid(cel.Formula, 2) & ", """")"
        End If
    Next cel

End Sub

Sub csvtoExcel()
'
' csvtoExcel Macro
'

'
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1)) _
        , TrailingMinusNumbers:=True
End Sub

Sub TextSplit1()
'
' summaryformat Macro
' summary dosyalarını doğru formatta çevir
'

'
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1)), DecimalSeparator:=".", ThousandsSeparator:=",", _
        TrailingMinusNumbers:=True
End Sub

Sub Change_to_Value()
'
' Change_to_Value Macro
' All Nulls to Values
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.NumberFormat = "0.00"
End Sub


Public Function stringToDbl(myString As Variant)
    
    'convert String to Double
    stringToDbl = CDbl(myString)

End Function



Sub Upload_GAD()
    Dim wb1 As Workbook


    Set wb1 = Application.ActiveWorkbook

    If WorksheetExists("GAD") = True Then
        wb1.Sheets("GAD").Delete
    End If
  Call ReadDataFromCloseFile("Bayi", "GAD")
       
       Sheets("SX WOW").Select
      Columns("J:BZ").Select
    Selection.Replace What:="#REF", Replacement:="GAD", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("Akış").Select
Range("A1").Select
  
End Sub

Sub ReadDataFromCloseFile(PrvShtNm As String, NShtNm As String)
    'On Error GoTo ErrHandler
     Application.DisplayAlerts = False
    'donot forget to restore the standard behavior at the end of your process:

   ' Application.ScreenUpdating = False
    Dim wb1 As Workbook, src As Workbook


    Set wb1 = Application.ActiveWorkbook
    
    Dim GetFileName As String, cellrange As String
        
  'Show the open dialog and pass the selected file name to the String variable "GetFileName"

    cellrange = CStr(NShtNm) & "B"

    GetFileName = Application.GetOpenFilename

    'They have cancelled.

    If GetFileName = "False" Then Exit Sub
    
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(GetFileName, True, True)
    
    ' GET THE TOTAL CBU ROWS FROM THE SOURCE WORKBOOK.
   
    ' COPY CBU DATA FROM SOURCE (CLOSE WORKGROUP) TO THE DESTINATION WORKBOOK.
  '  With wb1
  '      If WorksheetExists(ShtNm) = True Then
  '      wb1.Sheets(ShtNm).Delete
  '      End If
  '  End With
    
    src.Activate
    src.Sheets(PrvShtNm).Select
    src.Sheets(PrvShtNm).Copy Before:=wb1.Sheets("Akış")
    wb1.Sheets(PrvShtNm).Select
    wb1.Sheets(PrvShtNm).Name = NShtNm
    
    
    wb1.Activate
    Sheets("Akış").Select
    Range(cellrange).Select
    Range(cellrange).Value = GetFileName

   
   ' wbl.Range("'Akış'!" & cellrange).Select
   ' wbl.Range("'Akış'!" & cellrange).Value = "selam" 'GetFileName

    'src.Worksheets("CBU").Range("A:CZ").Copy
    'wb1.Worksheets("CBU").Range("A1").Select
    'wb1.ActiveSheet.Paste
   ' ThisWorkbook.Worksheets("EBU").Delete
    'Worksheets("Fixed").Delete
    'Worksheets("IS 15").Delete


    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
 
 Application.DisplayAlerts = True
   
'ErrHandler:
   ' Application.EnableEvents = True
  '  Application.ScreenUpdating = True
Call SortWorksheets
Call ListWorkSheetNamesNewWs

End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Sub ShowGUI()
UserForm1.Show
End Sub
Sub Upload_BYP()
    Dim wb1 As Workbook


    Set wb1 = Application.ActiveWorkbook

    If WorksheetExists("BYP") = True Then
        wb1.Sheets("BYP").Delete
    End If
  Call ReadDataFromCloseFile("Bayi Bilgileri", "BYP")
     Sheets("SX WOW").Select
    Range("A1").Select
   


   Columns("B:I").Select
    Selection.Replace What:="#REF", Replacement:="BYP", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("Akış").Select
Range("A1").Select

End Sub

Sub Upload_Invstr()
    Dim wb1 As Workbook


    Set wb1 = Application.ActiveWorkbook

    If WorksheetExists("Investor") = True Then
        wb1.Sheets("Investor").Delete
    End If
  Call ReadDataFromCloseFile("Yatirimci Bilgileri", "Investor")
       Sheets("SX WOW").Select
      Columns("J:BZ").Select
    Selection.Replace What:="#REF", Replacement:="Investor", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("Akış").Select
Range("A1").Select
   
    

End Sub
Sub Upload_HSD()
    Dim wb1 As Workbook


    Set wb1 = Application.ActiveWorkbook

    If WorksheetExists("HSD") = True Then
        wb1.Sheets("HSD").Delete
    End If
  Call ReadDataFromCloseFile("BAYI (1)", "HSD")
       Sheets("SX WOW").Select
      Columns("J:BZ").Select
    Selection.Replace What:="#REF", Replacement:="HSD", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("Akış").Select
Range("A1").Select
End Sub



Sub Upload_SHD()
    Dim wb1 As Workbook


    Set wb1 = Application.ActiveWorkbook

    If WorksheetExists("SHD") = True Then
        wb1.Sheets("SHD").Delete
    End If
  Call ReadDataFromCloseFile("Bayi", "SHD")
       Sheets("SX WOW").Select
      Columns("J:BZ").Select
    Selection.Replace What:="#REF", Replacement:="SHD", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("Akış").Select
Range("A1").Select
End Sub


Sub ListWorkSheetNamesNewWs()
'Updateby20140624
Dim xWs As Worksheet
On Error Resume Next
Application.DisplayAlerts = False
xTitleId = "Akış"
'Application.Sheets(xTitleId).Delete
'Application.Sheets.Add Application.Sheets(1)
Set xWs = Application.Sheet("Akış")
xWs.Name = xTitleId
For i = 2 To Application.Sheets.Count
    xWs.Range("AA" & (i - 1)) = Application.Sheets(i).Name
Next
Application.DisplayAlerts = True
End Sub

Sub SortWorksheets()
     
    Dim N As Integer
    Dim M As Integer
    Dim FirstWSToSort As Integer
    Dim LastWSToSort As Integer
    Dim SortDescending As Boolean
     
    SortDescending = False
     
    If ActiveWindow.SelectedSheets.Count = 1 Then
         
         'Change the 1 to the worksheet you want sorted first
        FirstWSToSort = 1
        LastWSToSort = Worksheets.Count
    Else
        With ActiveWindow.SelectedSheets
            For N = 2 To .Count
                If .Item(N - 1).Index <> .Item(N).Index - 1 Then
                    MsgBox "You cannot sort non-adjacent sheets"
                    Exit Sub
                End If
            Next N
            FirstWSToSort = .Item(1).Index
            LastWSToSort = .Item(.Count).Index
        End With
    End If
     
    For M = FirstWSToSort To LastWSToSort
        For N = M To LastWSToSort
            If SortDescending = True Then
                If UCase(Worksheets(N).Name) > UCase(Worksheets(M).Name) Then
                    Worksheets(N).Move Before:=Worksheets(M)
                End If
            Else
                If UCase(Worksheets(N).Name) < UCase(Worksheets(M).Name) Then
                    Worksheets(N).Move Before:=Worksheets(M)
                End If
            End If
        Next N
    Next M
     
End Sub
 
 
 Sub Adv_Filter_SX()
'
' Adv_Filter Macro
'
    Range("A10:BJ10000").Clear
    Range("BYP!A1:BJ20000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "SXEligibility!Criteria"), CopyToRange:=Range("A10:BJ10"), Unique:=False
        
        Call Tablo_Bul_SX
        Call Liste_aktar_SX
        
End Sub

Sub Adv_Filter_R()
'
' Adv_Filter Macro
'
    Range("A10:BJ10000").Clear
    Range("BYP!A1:BJ20000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "REligibility!Criteria"), CopyToRange:=Range("A10:BJ10"), Unique:=False
        
        Call Tablo_Bul_R
        Call Liste_aktar_R
        
End Sub

Sub Tablo_Bul_SX()
'
' Tablo_Bul Macro
'
    Application.Goto Reference:="R10C1"
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.Copy
    Range("F3").Select
    ActiveSheet.Paste Link:=True
End Sub
Sub Tablo_Bul_R()
'
' Tablo_Bul Macro
'
    Application.Goto Reference:="R10C1"
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.Copy
    Range("F3").Select
    ActiveSheet.Paste Link:=True
End Sub
Sub Liste_aktar_SX()
'
' Liste_aktar Macro
'
    Range("'SX WOW'!A2:A5000").Clear
    Range([INDIRECT("SXEligibility!j5")]).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "SXEligibility!j1"), CopyToRange:=Range("'SX WOW'!A2:A1500"), Unique:=False
End Sub
Sub Liste_aktar_R()
'
' Liste_aktar Macro
'
    Range("'R WOW'!A2:A5000").Clear
    Range([INDIRECT("REligibility!j5")]).AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "REligibility!j1"), CopyToRange:=Range("'R WOW'!A2:A2000"), Unique:=False
End Sub

Sub Adv_Filter_Inv()
'
' Adv_Filter Macro
'
    Range("A10:BJ5000").Clear
    Range("Investor!A1:V20000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range( _
        "InvEligibility!Criteria"), CopyToRange:=Range("A10:V10"), Unique:=False
        
     
        
End Sub

