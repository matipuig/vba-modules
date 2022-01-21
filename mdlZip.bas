Attribute VB_Name = "mdlZip"
Option Explicit

'''
' VBA-MODULES:
' Contains the necessary methods to create zip files, and insert and extract contents from them.
'''

'''
' Creates a zip file.
'''
Public Function createZipFile(ByVal zipPath As String) As Boolean
    If exists(zipPath) Then
        Err.Raise 1, , "File " & zipPath & " already exist!"
    End If
    Open zipPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Function

'''
' Inserts a content inside a zip file.
'''
Public Function insertIntoZip(ByVal contentPath As Variant, ByVal zipFilePath As Variant) As Boolean
    Dim strZipPath As String: strZipPath = Conversion.CStr(zipFilePath)
    throwErrorIfZipNotExist strZipPath
    Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
    Dim previousItemsCount As Integer: previousItemsCount = ShellApp.Namespace(zipFilePath).items.Count
    ShellApp.Namespace(zipFilePath).CopyHere contentPath
    On Error Resume Next
    Do Until ShellApp.Namespace(zipFilePath).items.Count = previousItemsCount + 1
        DoEvents
    Loop
    On Error GoTo 0
    insertIntoZip = True
End Function

'''
' Unzips the content into a directory.
'''
Public Function extractInto(ByVal zipFilePath As Variant, ByVal targetDirectory As Variant) As Boolean
    Dim strZipPath As String: strZipPath = Conversion.CStr(zipFilePath)
    throwErrorIfZipNotExist strZipPath
    If exists(targetDirectory) = False Then
        Err.Raise 1, , "Directory " & targetDirectory & " does not exist."
    End If
    Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(targetDirectory).CopyHere ShellApp.Namespace(zipFilePath).items, vbHide
    extractInto = True
End Function

'''
' 
' PRIVATE METHODS
'
'''

'''
' Throws an error if the zip file does not exist.
'''
Private Function throwErrorIfZipNotExist(ByVal zipPath As String) As Boolean
    If exists(zipPath) = False Then
        Err.Raise 1, , "File " & zipPath & " does not exist"
    End If
    throwErrorIfZipNotExist = False
End Function

'''
' Returns if a file exists.
'''
Private Function exists(ByVal filePath As String) As Boolean
    If Len(Dir(filePath)) <> 0 Then
        exists = True
    End If
End Function
