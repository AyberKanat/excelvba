Sub Tablo_Bul()
'
' Tablo_Bul Macro
'
    Application.Goto Reference:="R10C1"
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.Copy
    Range("F3").Select
    ActiveSheet.Paste Link:=True