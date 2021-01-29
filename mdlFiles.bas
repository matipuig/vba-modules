Attribute VB_Name = "mdlFiles"
Option Explicit
'''
'   FILES METHODS.
'   This module lets you interact with files in VBA.
'''

Const DIALOG_TYPE_FOLDER As Single = 1
Const DIALOG_TYPE_FILE As Single = 2

'''
'   Gets FileSystemObject instance.
'''
Public Function getFSO() As Object
    Static FSO As Object
    If FSO Is Nothing Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set getFSO = FSO
End Function

'''
'   Show the OpenFileDialog modal box.
'   Returns the specifed folder or "".
'''
Public Function chooseFolder(Optional ByVal title As String = "Por favor, elige una Dir") As String
    chooseFolder = openDialog(DIALOG_TYPE_FOLDER, title)
End Function

'''
'   Show the OpenFileDialog modal box.
'   Returns the specifed file or "".
'''
Public Function chooseFile(Optional ByVal title As String = "Por favor, elige un archivo") As String
    chooseFile = openDialog(DIALOG_TYPE_FILE, title)
End Function

'''
'   Returns if the folder exists.
'''
Public Function folderExists(ByVal folder As String) As Boolean
    Dim FSO As Object: Set FSO = getFSO
    folderExists = FSO.folderExists(folder)
    Set FSO = Nothing
End Function

'''
'   Returns if the file exists.
'''
Public Function fileExists(ByVal file As String) As Boolean
    Dim FSO As Object: Set FSO = getFSO
    fileExists = FSO.fileExists(file)
    Set FSO = Nothing
End Function

'''
'   Returns an array with all files in folder.
'''
Public Function getFiles(ByVal dirPath As String, Optional ByVal withSubdirs As Boolean = False)
    Dim FSO As Object: Set FSO = getFSO
    Dim srcDir As Object: Set srcDir = FSO.getFolder(dirPath)
    Dim allFiles As Object: Set allFiles = CreateObject("Scripting.Dictionary")
    Dim file As Object
    Dim subdir As Object
    Dim subdirFiles
    Dim subdirFile

    For Each file In srcDir.Files
        allFiles.Add file.Path, ""
    Next file
    
    If withSubdirs = True Then
        For Each subdir In srcDir.Subfolders
            subdirFiles = getFiles(subdir.Path, True)
            If IsEmpty(subdirFiles) Then GoTo nextSubdir
            For Each subdirFile In subdirFiles
                allFiles.Add subdirFile, ""
            Next subdirFile
            DoEvents
nextSubdir:
        Next subdir
    End If

    Set file = Nothing
    Set subdir = Nothing
    Set srcDir = Nothing
    Set FSO = Nothing
    getFiles = allFiles.keys
    Set allFiles = Nothing
End Function

'''
'   Gets the filename with the extension.
'''
Public Function getFileName(ByVal filePath As String) As String
    getFileName = ""
    If filePath = "" Then Exit Function
    Dim arrName: arrName = Strings.Split(filePath, "\")
    getFileName = arrName(UBound(arrName))
End Function

'''
'   Gets the filename extension.
'   Returns "" if it cannot find it.
'''
Public Function getExtension(ByVal filePath As String) As String
    getExtension = ""
    Dim fileName As String: fileName = getFileName(filePath)
    Dim arrName: arrName = Strings.Split(fileName, ".")
    If Information.IsEmpty(arrName) Then Exit Function
    Dim lastIndex: lastIndex = UBound(arrName)
    If lastIndex = 0 Then Exit Function
    Dim extension As String: extension = arrName(lastIndex)
    getExtension = Strings.LCase(extension)
End Function

'''
'   Deletes a file.
'''
Public Function deleteFile(ByVal filePath As String) As Boolean
   If fileExists(filePath) Then
      SetAttr filePath, vbNormal
      Kill filePath
   End If
   deleteFile = True
End Function

'''
'   Converts a file to Base64.
'''
Public Function toBase64(ByVal filePath As String) As String
    If fileExists(filePath) = False Then
        Err.Raise 1, , "File ''" & filePath & "'' doesn't exist."
    End If

    Const UseBinaryStreamType = 1
    Dim streamInput: Set streamInput = getADODBStream
    Dim xmlDoc: Set xmlDoc = getXMLDOM
    Dim xmlElem: Set xmlElem = xmlDoc.createElement("tmp")
    
    streamInput.Open
    streamInput.Type = UseBinaryStreamType
    streamInput.LoadFromFile filePath
    xmlElem.DataType = "bin.base64"
    xmlElem.nodeTypedValue = streamInput.Read
    toBase64 = replace(xmlElem.Text, vbLf, "")
    streamInput.Close

    Set streamInput = Nothing
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
End Function


'''
'
'   PRIVATE METHODS
'
'''

'''
'   Show the OpenFileDialog modal box.
'   Returns the specifed fileor "".
'''
Private Function openDialog(ByVal dialogType As Single, ByVal title As String) As String
    Dim dialogBox As FileDialog
    Dim selection As String: selection = ""
        
    If dialogType = DIALOG_TYPE_FOLDER Then
        Set dialogBox = Application.FileDialog(msoFileDialogFolderPicker)
    ElseIf dialogType = DIALOG_TYPE_FILE Then
        Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
    End If
    
    With dialogBox
        .title = title
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo EndFunction
        selection = .SelectedItems(1)
    End With

EndFunction:
    openDialog = selection
    Set dialogBox = Nothing
End Function

'''
'   Gets the ADODBStream instance.
'''
Private Function getADODBStream() As Object
    Static AdoDBStream As Object
    If AdoDBStream Is Nothing Then
        Set AdoDBStream = CreateObject("ADODB.Stream")
    End If
    Set getADODBStream = AdoDBStream
End Function

'''
'   Gets the XMLDOM object.
'''
Private Function getXMLDOM() As Object
    Static XMLDOM As Object
    If XMLDOM Is Nothing Then
        Set XMLDOM = CreateObject("Microsoft.XMLDOM")
    End If
    Set getXMLDOM = XMLDOM
End Function




