 Sub Upload_BYP()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("BYP") = True Then
        wb1.Sheets("BYP").Delete
    End If
  Call ReadDataFromCloseFile("Bayi Bilgileri", "BYP")
     Sheets("StartUp").Select
    Range("A1").Select
End Sub


 Sub Upload_GAD()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("GAD") = True Then
        wb1.Sheets("GAD").Delete
    End If
  Call ReadDataFromCloseFile("Bayi", "GAD")
     Sheets("StartUp").Select
    Range("A1").Select
End Sub

 Sub Upload_HSD()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("HSD") = True Then
        wb1.Sheets("HSD").Delete
    End If
  Call ReadDataFromCloseFile("BAYI (1)", "HSD")
     Sheets("StartUp").Select
    Range("A1").Select
End Sub

 Sub Upload_SHD()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("SHD") = True Then
        wb1.Sheets("SHD").Delete
    End If
  Call ReadDataFromCloseFile("Bayi", "SHD")
     Sheets("StartUp").Select
    Range("A1").Select
End Sub


 Sub Upload_SGM()
    Dim wb1 As Workbook
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("SGM") = True Then
        wb1.Sheets("SGM").Delete
    End If
  Call ReadDataFromCloseFile("Segment", "SGM")
    Sheets("StartUp").Select
    Range("A1").Select
End Sub

 Sub Upload_TRG()
    Dim wb1 As Workbook
    Dim MySheet As String
    Set wb1 = Application.ActiveWorkbook
    If WorksheetExists("TRG") = True Then
        wb1.Sheets("TRG").Delete
    End If
  Call ReadDataFromCloseFile("xx", "TRG")
    
    MySheet = wb1.ActiveSheet.Name
    wb1.Sheets(MySheet).Name = "TRG"

    ' Debug.Print MySheet
    
    Sheets("StartUp").Select
    Range("A1").Select
End Sub
