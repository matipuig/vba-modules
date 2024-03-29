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
' Creates a folder if it doesn't exist.
'''
Public Function makeDir(ByVal folderPath As String) As Boolean
    If folderExists(folderPath) Then
        makeDir = True
        Exit Function
    End If
    Dim FSO As Object: Set FSO = getFSO
    FSO.createfolder folderPath
    makeDir = True
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
'   Returns an array with all folders in folder.
'''
Public Function getFolders(ByVal dirPath As String, Optional ByVal withSubdirs As Boolean = False)
    Dim FSO As Object: Set FSO = getFSO
    Dim srcDir As Object: Set srcDir = FSO.getFolder(dirPath)
    Dim allDirs As Object: Set allDirs = CreateObject("Scripting.Dictionary")
    Dim subdir As Object
    Dim subdirDirs
    Dim subdirDir
   
    For Each subdir In srcDir.subfolders
        allDirs.Add subdir, ""
        If withSubdirs = True Then
            subdirDirs = getFolders(subdir.Path, True)
            If Information.IsEmpty(subdirDirs) Then GoTo nextSubdir
            For Each subdirDir In subdirDirs
                allDirs.Add subdirDir, ""
            Next subdirDir
        End If
nextSubdir:
        DoEvents
    Next subdir
    
    getFolders = allDirs.keys
    Set subdir = Nothing
    Set srcDir = Nothing
    Set FSO = Nothing
    Set allDirs = Nothing
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
        DoEvents
    Next file
    
    If withSubdirs = True Then
        For Each subdir In srcDir.Subfolders
            subdirFiles = getFiles(subdir.Path, True)
            If Information.IsEmpty(subdirFiles) Then GoTo nextSubdir
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
' Saves a text content into a file.
'''
Public Function saveTextFile(ByVal filePath As String, ByVal content As String, Optional charset As String = "utf-8") As Boolean
    If fileExists(filePath) Then
        Err.Raise 1, , "File " & filePath & " already exists."
    End If
    
    Dim ADODB As Object: Set ADODB = CreateObject("ADODB.Stream")
    ADODB.Type = 2 'Write.
    ADODB.charset = charset
    ADODB.Open
    ADODB.WriteText content
    ADODB.SaveToFile filePath, 2
    Set ADODB = Nothing
    saveTextFile = True
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




