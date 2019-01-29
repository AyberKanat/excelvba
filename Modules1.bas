Sub Upload_BYP()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("01-Bayi Bilgileri") = True Then
        wb1.Sheets("01-Bayi Bilgileri").Delete
    End If
  
 Call ReadDataFromCloseFile("Bayi Bilgileri", "01-Bayi Bilgileri")
 Call SortWorksheets
     Sheets("Main").Select
    Range("A1").Select
End Sub

Sub Upload_Investor()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("02-Yatırımcı Bilgileri") = True Then
        wb1.Sheets("02-Yatırımcı Bilgileri").Delete
    End If
  
 Call ReadDataFromCloseFile("Yatirimci Bilgileri", "02-Yatırımcı Bilgileri")
 Call SortWorksheets
 
     Sheets("Main").Select
    Range("A1").Select
End Sub


Sub ReadDataFromCloseFile(PrvShtNm As String, NShtNm As String)
    'On Error GoTo ErrHandler
     Application.DisplayAlerts = False
    'do not forget to restore the standard behavior at the end of your process:

   ' Application.ScreenUpdating = False
    Dim wb1 As Workbook, src As Workbook


    Set wb1 = Application.ActiveWorkbook
    
    Dim GetFileName As String
        
  'Show the open dialog and pass the selected file name to the String variable "GetFileName"


    GetFileName = Application.GetOpenFilename

    'They have cancelled.

    If GetFileName = "False" Then Exit Sub
    
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(GetFileName, True, True)
    
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK.
   
    
    src.Activate
    src.Sheets(PrvShtNm).Select
    src.Sheets(PrvShtNm).Copy Before:=wb1.Sheets("Main")
    wb1.Sheets(PrvShtNm).Select
    wb1.Sheets(PrvShtNm).Name = NShtNm
    
    
   
 
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
 
 Application.DisplayAlerts = True
   
'ErrHandler:
   ' Application.EnableEvents = True
  '  Application.ScreenUpdating = True


End Sub

Sub Upload_HSD()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("03-HSD") = True Then
        wb1.Sheets("03-HSD").Delete
    End If
  
 Call ReadDataFromCloseFile("BAYI (1)", "03-HSD")
 Call SortWorksheets
 
     Sheets("Main").Select
    Range("A1").Select
End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

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

