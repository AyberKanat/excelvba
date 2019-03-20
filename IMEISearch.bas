Private Sub IMEI_Search()
Dim DOK As String
Dim FeedBack As Integer

DOK = InputBox("Input IMEI number:?")
For i = 1 To Sheets.Count
  
      For X = 1 To 200000
       If Sheets(i).Cells(X, 7) = DOK Then
FeedBack = MsgBox("IMEI: " & DOK & " is on sheets: " & Sheets(i).Name & vbNewLine & " Open Sheet?", vbYesNoCancel)
        Select Case FeedBack
        Case 6
            Sheets(i).Activate
            ActiveSheet.Cells(X, 1).Select
            GoTo 20
        Case 7
            GoTo 20
        Case 2
            GoTo 10
        End Select
       End If
20    Next X
      
Next i
10
End Sub
