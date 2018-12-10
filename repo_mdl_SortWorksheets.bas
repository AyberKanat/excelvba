Option Explicit
 
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