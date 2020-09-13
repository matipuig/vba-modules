Attribute VB_Name = "mdlStrings"
Option Explicit
'''
'   STRINGS.
'   Contains some functions used often for strings.
'''

'''
'   Returns if the source text has the subtext.
'''
Public Function hasSubtext(ByVal originaltext As String, ByVal searchedSubtext As String, Optional ByVal matchCase As Boolean = False) As Boolean
    If matchCase = False Then
        originaltext = Strings.LCase(originaltext)
        searchedSubtext = Strings.LCase(searchedSubtext)
    End If
    hasSubtext = Strings.InStr(originaltext, searchedSubtext) > 0
End Function

'''
'   Get RegExp object.
'''
Public Function getRegExp(ByVal pattern As String, ByVal ignoreCase As Boolean, ByVal globalRegex As Boolean, ByVal multiLine As Boolean) As Object
    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    RegExp.Global = globalRegex
    RegExp.ignoreCase = ignoreCase
    RegExp.multiLine = multiLine
    Set getRegExp = RegExp
End Function

'''
'   Tests a regex in a string.
'''
Public Function testRegex(ByVal originaltext As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    Dim RegExp As Object: Set RegExp = getRegExp(pattern, ignoreCase, True, True)
    testRegex = RegExp.test(originaltext)
End Function

'''
'   Executes one replace in a string using regex.
'''
Public Function replaceOneWithRegex(ByVal originaltext As String, ByVal searchedPattern As String, ByVal replacement As String, Optional ByVal ignoreCase As Boolean = False) As String
    Dim RegExp As Object: Set RegExp = getRegExp(searchedPattern, ignoreCase, False, True)
    replaceOneWithRegex = RegExp.Replace(originaltext, replacement)
End Function

'''
'   Executes all replacements in a string using regex.
'''
Public Function replaceWithRegex(ByVal originaltext As String, ByVal searchedPattern As String, ByVal replacement As String, Optional ByVal ignoreCase As Boolean = False) As String
    Dim RegExp As Object: Set RegExp = getRegExp(searchedPattern, ignoreCase, True, True)
    replaceWithRegex = RegExp.Replace(originaltext, replacement)
End Function

'''
'   Returns all the matches found for a regex in the original text as String().
'   It returns FALSE if it cannot find any.
'''
Public Function executeRegex(ByVal originaltext As String, ByVal pattern As String, Optional ByVal ignoreCase As Boolean = False, Optional ByVal globalRegex As Boolean = True, Optional ByVal multiLine As Boolean = True)
    Dim RegExp As Object: Set RegExp = getRegExp(pattern, ignoreCase, globalRegex, multiLine)
    Dim allMatches: Set allMatches = RegExp.Execute(originaltext)
    
    If allMatches.Count = 0 Then
        executeRegex = False
        Exit Function
    End If
    
    Dim I As Long, J As Long, result As String
    Dim SEPARATOR As String: SEPARATOR = "-|-SEP-|-"
    For I = 0 To allMatches.Count - 1
        result = result & allMatches.Item(I).Value & SEPARATOR
        For J = 0 To allMatches.Item(I).submatches.Count - 1
            result = result & allMatches.Item(I).submatches.Item(J) & SEPARATOR
        Next
    Next
        
    result = Strings.Left(result, Strings.Len(result) - Strings.Len(SEPARATOR))
    executeRegex = Strings.Split(result, SEPARATOR)
End Function
