VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub prueba()

Dim fList As String
Dim fName As String
Dim sExt As String
Dim sFil As String
sFil = ActiveWorkbook.Name
sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
fName = Dir("C:\Users\Apollo\Desktop\Excel Tests\Backup Delete Folder\*" & sExt)
' The variable fName now contains the name of the files within the specified path.
Do While fName <> ""
' Store the current file in the string fList.
fList = fList & vbNewLine & fName
' Get the next files in the specified path.
fName = Dir()
' The variable fName now contains the name of the next files in the specified path.
Loop
' Display the list of files in a message box.
MsgBox ("List of Files:" & fList)

End Sub

Sub FirstFile()

Dim NameFile As String
Dim sNombre As String
Dim sExt As String
Dim sFil As String
Dim sFile As String
Dim myFileSystemObject As FileSystemObject
sFil = ActiveWorkbook.Name
sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
sNombre = "C:\Users\Apollo\Desktop\Excel Tests\Backup " & sFile & "\*" & sExt
NameFile = Dir("C:\Users\Apollo\Desktop\Excel Tests\Backup " & sFile & "\*" & sExt)

MsgBox NameFile


End Sub

Sub test2()

Dim fso
Dim carpet
Dim files
Dim file
Dim count
Dim NameFile As String
Dim sNombre As String
Dim sExt As String
Dim sFil As String
Dim sFile As String
Dim sDeleteFile As String
Dim myFileSystemObject As FileSystemObject
sFil = ActiveWorkbook.Name
sFile = VBA.Mid(sFil, 1, VBA.InStr(sFil, ".x") - 1)
sExt = VBA.Right(sFil, Len(sFil) - WorksheetFunction.Find(".", sFil) + 1)
sNombre = "C:\Users\Apollo\Desktop\Excel Tests\Backup " & sFile & "\*" & sExt
NameFile = Dir("C:\Users\Apollo\Desktop\Excel Tests\Backup " & sFile & "\*" & sExt)
ChDir "C:\Users\Apollo\Desktop\Excel Tests\Backup Delete Folder\"
Set fso = CreateObject("scripting.filesystemobject")
Set carpet = fso.GetFolder(CurDir())
Set files = carpet.files
For Each file In files
count = count + 1
Next
If count > 10 Then
    MsgBox NameFile
Else
    MsgBox "There's 10 or less"
End If

sDeleteFile = NameFile.Name
MsgBox sDeleteFile
End Sub


