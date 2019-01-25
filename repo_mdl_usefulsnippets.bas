'Snippets Page
'File exists

Sub FileExists()
Dim fso
Dim strFile As String
strFile = “C:\Test.xls” ‘ change to match the file w/Path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If Not fso.FileExists(strFile) Then
MsgBox strFile & ” was not located.”, vbInformation, “File Not Found”
Else
MsgBox strFile & ” has been located.”, vbInformation, “File Found”
End If
End Sub

'Copy File

Sub CopyFile()
Dim fso
Dim strFile As String,
Dim strSrcFol As String
Dim strDestFol As String
strFfile = “test.xls” ‘ change to match the file name
strSrcFol = “C:\” ‘ change to match the source folder path
strDestFol = “E:\” ‘ change to match the destination folder path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If Not fso.FileExists(strSrcfol & strFile) Then
MsgBox strSrcFol & strFile & ” does not exist!”, vbExclamation, “Source File Missing”
ElseIf Not fso.FileExists(strDestFol & strFile) Then
fso.CopyFile (strSrcFol & strFile), strDestFol, True
Else
MsgBox strDestFol & strFile & ” already exists!”, vbExclamation, “Destination File Exists”
End If
End Sub

'Move a File

Sub MoveFile()
Dim fso
Dim strFile As String
Dim strSrcFol As String
Dim strDestfol As String
strFile = “test.xls” ‘ change to match the file name
strSrcFol = “C:\” ‘ change to match the source folder path
strDestFol = “E:\” ‘ change to match the destination folder path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If Not fso.FileExists(strSrcFol & strFile) Then
MsgBox strSrcFol & strFile & ” does not exist!”, vbExclamation, “Source File Missing”
ElseIf Not fso.FileExists(strDestFol & strFile) Then
fso.MoveFile (strSrcFol & strFile), strDestFol
Else
MsgBox strDestFol & strFile & ” already exists!”, vbExclamation, “Destination File Exists”
End If
End Sub

'Folder Exists

Sub FolderExists()
Dim fso
Dim strFolder As String
folder = “C:\My Documents” ‘ change to match the folder path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If fso.FolderExists(strFolder) Then
MsgBox strFolder & ” is a valid folder/path.”, vbInformation, “Path Exists”
Else
MsgBox strFolder & ” is not a valid folder/path.”, vbInformation, “Invalid Path”
End If
End Sub

'Create Folder

Sub CreateFolder()
Dim fso
Dim strFol As String
strFol = “c:\MyFolder” ‘ change to match the folder path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If Not fso.FolderExists(strFol) Then
fso.CreateFolder (strFol)
Else
MsgBox strFol & ” already exists!”, vbExclamation, “Folder Exists”
End If
End Sub

'Copy Folder

Sub CopyFolder()
Dim fso
Dim strSrcFol As String
Dim strDestdfol As String
strSrcfol = “c:\MyFolder” ‘ change to match the source folder path
strDestfol = “e:\MyFolder” ‘ change to match the destination folder path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If Not fso.FolderExists(strDestFol) Then
fso.CopyFolder strSrcFol, strDestFol
Else
MsgBox strDestFol & ” already exists!”, vbExclamation, “Folder Exists”
End If
End Sub 

' Move Folder – Only if allowed otherwise it will error out (so check if you have permissions)

Sub MoveFolder()
Dim fso
Dim strSrcFol As String
dim strDestFol As String
strSrcFol = “c:\MyFolder” ‘ change to match the source folder path
strDestFol = “e:\MyFolder” ‘ change to match the destination folder path
Set fso = CreateObject(“Scripting.FileSystemObject”)
If Not fso.FolderExists(strDestFol) Then
fso.MoveFolder strSrcFol, strDestFol
Else
MsgBox strDestFol & ” already exists!”, vbExclamation, “Folder Exists”
End If
End Sub
