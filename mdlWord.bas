Attribute VB_Name = "mdlWord"
Option Explicit

'''
'   MICROSOFT WORD METHODS.
'   This module contains the methods needed for microsoft word.
'''

'''
'   Opens a Word App and return its reference.
'''
Private Function obtenerWordApp(Optional ByVal visible As Boolean = False) As Object
    Static wordApp As Object
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
        wordApp.visible = visible
    End If
    Set obtenerWordApp = wordApp
End Function

'''
'   Reads the complete Word Content. Can close it at the end.
'''
Public Function getContent(ByVal filePath As String, Optional ByVal closeAtEnding = False) As String
    Dim wordApp As Object: Set wordApp = obtenerWordApp
    Dim wordDoc As Object: Set wordDoc = wordApp.documents.Open(filePath)
    getContent = wordDoc.Range.text
    wordDoc.Close
    Set wordDoc = Nothing
    If closeAtEnding = True Then closeApp
End Function

'''
'   Closes the word app.
'''
Public Function closeApp() As Boolean
    Dim wordApp As Object: Set wordApp = obtenerWordApp
    wordApp.Quit
    Set wordApp = Nothing
    closeApp = True
End Function

