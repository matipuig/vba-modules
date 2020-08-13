Attribute VB_Name = "mdlWord"
Option Explicit

'''
'   MICROSOFT WORD METHODS.
'   This module contains the methods needed for microsoft word.
'''
Const REPLACE_TYPE_FIRST As Single = 1
Const REPLACE_TYPE_ALL As Single = 2

'''
'   Opens a Word App and return its reference.
'''
Public Function getWordApp(Optional ByVal visible As Boolean = False) As Object
    Static wordApp As Object
    If isApplication(wordApp) = False Then
        Set wordApp = CreateObject("Word.Application")
        wordApp.visible = visible
    End If
    Set getWordApp = wordApp
End Function

'''
'   Reads the complete Word Content. Can close it at the end.
'''
Public Function getContent(ByVal filePath As String, Optional ByVal closeAppAtEnding = False) As String
    Dim wordApp As Object: Set wordApp = getWordApp
    Dim wordDoc As Object: Set wordDoc = wordApp.documents.Open(filePath)
    getContent = wordDoc.range.Text
    wordDoc.Close
    Set wordDoc = Nothing
    If closeAppAtEnding = True Then closeApp
    Set wordAPP = Nothing
End Function

'''
'   Replaces all ocurrences in range in word.
'''
Public Function replaceAllOcurrences(ByRef range As Object, ByVal replace As String, ByVal replacement As String, Optional ByVal matchCase As Boolean = True) As Boolean
    replaceAllOcurrences = replaceInRange(REPLACE_TYPE_ALL, range, replace, replacement, matchCase)
End Function

'''
'   Replaces first ocurrence in range in word.
'''
Public Function replaceFirstOcurrence(ByRef range As Object, ByVal replace As String, ByVal replacement As String, Optional ByVal matchCase As Boolean = True) As Boolean
    replaceFirstOcurrence = replaceInRange(REPLACE_TYPE_FIRST, range, replace, replacement, matchCase)
End Function

'''
'   Closes the word app.
'''
Public Function closeApp() As Boolean
    Dim wordApp As Object: Set wordApp = getWordApp
    If isApplication(wordApp) = False Then
        closeApp = True
        Exit Function
    End If
    
    Dim doc
    For Each doc In wordApp.documents
        doc.Close SaveChange:=0
    Next doc
    wordApp.Quit
    Set wordApp = Nothing
    closeApp = True
End Function

'
'
'   PRIVATE METHODS
'
'

'''
'   Checks if is open.
'''
Private Function isApplication(ByRef object) As Boolean
    isApplication = False
    If Information.TypeName(object) = "Application" Then
        isApplication = True
    End If
End Function

'''
'   Replaces the ocurrences in word.
'''
Private Function replaceInRange(ByVal replaceType As Integer, ByRef range As Object, ByVal replace As String, ByVal replacement As String, Optional ByVal matchCase As Boolean = True) As Boolean
    range.Find.ClearFormatting
    range.Find.replacement.ClearFormatting
    With range.Find
        .Text = replace
        .replacement.Text = replacement
        .Forward = True
        .Wrap = 0
        .Format = False
        .matchCase = matchCase
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If replaceType = REPLACE_TYPE_FIRST Then
        range.Find.Execute replace:=1
    ElseIf replaceType = REPLACE_TYPE_ALL Then
        range.Find.Execute replace:=2
    End If
End Function
