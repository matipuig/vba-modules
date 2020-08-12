Attribute VB_Name = "mdlArrays"
Option Explicit
'''
'   ARRAYS.
'   Contains all functions with arrays.
'''

'''
'   Add some content to the array.
'   If it's not an array or the array is empty, returns a new array with only this new element.
'''
Public Function addToArray(ByRef sourceArray, ByVal newValue) As Variant
    Dim newArray(0)
    If Information.IsArray(sourceArray) = False Then
        newArray(0) = newValue
        addToArray = newArray
        Exit Function
    ElseIf Information.IsEmpty(sourceArray) Then
        newArray(0) = newValue
        addToArray = newArray
        Exit Function
    End If
       
    Dim arraySize As Single: arraySize = UBound(sourceArray) + 1
    ReDim Preserve sourceArray(arraySize)
    sourceArray(arraySize) = newValue
    addToArray = sourceArray
End Function

'''
'   Join two arrays and return the result.
'''
Public Function joinArrays(ByRef firstArray, ByRef secondArray)
    Dim jointArrays() As Variant
    Dim len1 As Single: len1 = UBound(firstArray)
    Dim len2 As Single: len2 = UBound(secondArray)
    Dim lenRe As Single: lenRe = len1 + len2 + 1
    Dim counter As Single: counter = 0
    ReDim jointArrays(0 To lenRe)

    For counter = 0 To len1
        jointArrays(counter) = firstArray(counter)
    Next
    For counter = 0 To len2
        jointArrays(counter + len1 + 1) = secondArray(counter)
    Next
    joinArrays = jointArrays
End Function

'''
'   Returns if it found the specified value.
'''
Public Function isInArray(ByVal sourceArray, ByVal searchedValue) As Boolean
    isInArray = False
    Dim element
    For Each element In sourceArray
        If element = searchedValue Then
            isInArray = True
            Exit Function
        End If
    Next element
End Function

''
'  Searches with InStr inside an array and returns the first result.
'  If it doesn't find anything, returns "".
''
Public Function searchInArray(ByVal sourceArray, ByVal searchedValue As String, Optional ByVal matchCase As Boolean = False) As String
    searchInArray = ""
    Dim element
    Dim tmpElement As String
    If matchCase = False Then
        searchedValue = Strings.LCase(searchedValue)
    End If
    
    For Each element In sourceArray
        element = Conversion.CStr(element)
        If matchCase = False Then
            tmpElement = Strings.LCase(element)
        Else
            tmpElement = element
        End If
        If Strings.InStr(tmpElement, searchedValue) > 0 Then
            searchInArray = element
            Exit Function
        End If
        DoEvents
    Next element
End Function