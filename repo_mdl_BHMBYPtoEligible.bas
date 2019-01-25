Option Explicit

Sub Adv_Filter()
'
' Adv_Filter Macro
'
    Sheets("Eligibility").Select
    Range("A1").Select
    Range("A1:BI5000").Clear
    Range("'BYP'!A1:BI20000").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("Scheme!BHMCriteria"), CopyToRange:=Range("A1:BI1"), Unique:=False
Call Count_Table
End Sub

Sub Count_Table()
    Dim DCount As Integer
    Application.Goto Reference:="R10C1"
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    DCount = ActiveCell.Row - 1
'Debug.Print DCount
Call Eligibility_Load(DCount)
End Sub

Sub Eligibility_Load(Counter As Integer)

Dim objD As CBHMDealers
Dim BHMDealers As Collection
Dim lDealerIDCntr As Long

Set BHMDealers = New Collection

Sheets("Eligibility").Select
Range("M1").Activate

For lDealerIDCntr = 1 To Counter
    
    Set objD = New CBHMDealers
    BHMDealers.Add objD
    
    BHMDealers(lDealerIDCntr).Sequence = lDealerIDCntr
    BHMDealers(lDealerIDCntr).Code = ActiveCell.Offset(lDealerIDCntr, 0)
    BHMDealers(lDealerIDCntr).Title = ActiveCell.Offset(lDealerIDCntr, 4)
    BHMDealers(lDealerIDCntr).SExcName = ActiveCell.Offset(lDealerIDCntr, 14)
    BHMDealers(lDealerIDCntr).RgnID = ActiveCell.Offset(lDealerIDCntr, 9)
    BHMDealers(lDealerIDCntr).City = ActiveCell.Offset(lDealerIDCntr, 10)
    BHMDealers(lDealerIDCntr).StartTime = ActiveCell.Offset(lDealerIDCntr, 16)
    
    ' Debug.Print BHMDealers(lDealerIDCntr).Code
    ' Debug.Print BHMDealers(lDealerIDCntr).Title
    ' Debug.Print BHMDealers(lDealerIDCntr).SExcName
    ' Debug.Print BHMDealers(lDealerIDCntr).RgnID
    ' Debug.Print BHMDealers(lDealerIDCntr).City
    ' Debug.Print BHMDealers(lDealerIDCntr).StartTime


Next lDealerIDCntr

Sheets("Calculation").Select
Range("A2").Activate


For Each objD In BHMDealers

    ActiveCell.Offset(objD.Sequence, 0).Value = objD.Sequence
    ActiveCell.Offset(objD.Sequence, 1).Value = objD.Code
    ActiveCell.Offset(objD.Sequence, 2).Value = objD.Title
    ActiveCell.Offset(objD.Sequence, 3).Value = objD.SExcName
    ActiveCell.Offset(objD.Sequence, 4).Value = objD.RgnID
    ActiveCell.Offset(objD.Sequence, 5).Value = objD.City
    ActiveCell.Offset(objD.Sequence, 6).Value = objD.StartTime

Next

'Print out values of each CustID in the collection

' For Each objD In BHMDealers


' Next

' 'Set objects to Nothing to delete them
' Set BHMDealers = Nothing
End Sub

