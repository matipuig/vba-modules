Attribute VB_Name = "mdlJSON"

Option Explicit
Const QUOTES As String = """"
Const ESCAPED_QUOTES As String = "\" & QUOTES

'''
'   Escapes the string to return
'''
Private Function escapeString(ByVal text As String) As String
    Dim escapedString As String: escapedString = Replace(text, QUOTES, ESCAPED_QUOTES)
    escapeString = QUOTES & escapedString & QUOTES
End Function

'''
'   Converts simple types (string, number, null and boolean) to JSON.
'''
Private Function convertSimpleTypeToJSON(ByRef something As Variant) As String
    If Information.VarType(something) = vbString Then
        convertSimpleTypeToJSON = escapeString(something)
        Exit Function
    
    ElseIf something = True Then
        convertSimpleTypeToJSON = "true"
        Exit Function
    ElseIf something = False Then
        convertSimpleTypeToJSON = "false"
        Exit Function
    
    ElseIf Information.IsNumeric(something) Then
        convertSimpleTypeToJSON = something
        Exit Function
        
    ElseIf Information.IsNull(something) Then
        convertSimpleTypeToJSON = "null"
        Exit Function
    End If
    Err.Raise "Type not accepted for JSON."
End Function

'''
'   Converts an array to JSON.
'''
Private Function arrayToJSON(ByRef someArray As Variant) As String
    Dim I As Double
    Dim tmpConversion As String
    Dim arraySize As Single: arraySize = UBound(someArray)
    
    If arraySize = 0 Then
        arrayToJSON = "[]"
        Exit Function
    End If
    For I = LBound(someArray) To arraySize
        tmpConversion = convertToJSON(someArray(I))
        arrayToJSON = arrayToJSON & tmpConversion
        If I < arraySize Then
            arrayToJSON = arrayToJSON & ","
        End If
    Next I
    arrayToJSON = "[" & arrayToJSON & "]"
End Function

'''
'   Converts a Dictionary to JSON.
'''
Private Function dictionaryToJSON(ByRef something As Variant) As String
    Dim I As Integer
    Dim tmpKey As String
    Dim tmpValue As String
    Dim tmpKeyValue As String
    Dim dictionary As Object: Set dictionary = something
    Dim dictionaryCount As Single: dictionaryCount = dictionary.Count - 1
    For I = 0 To dictionaryCount
        tmpKey = escapeString(dictionary.keys()(I))
        tmpValue = convertToJSON(dictionary.items()(I))
        tmpKeyValue = tmpKey & ":" & tmpValue
        If I < dictionaryCount Then
            tmpKeyValue = tmpKeyValue & ","
        End If
        dictionaryToJSON = dictionaryToJSON & tmpKeyValue
    Next I
    dictionaryToJSON = "{" & dictionaryToJSON & "}"
End Function

'''
'   Converts the specified received thing in JSON string.
'''
Public Function convertToJSON(ByRef something As Variant) As String
    Dim someType As Variant: someType = Information.VarType(something)
    
    If someType = vbBoolean Or someType = vbString Or Information.IsNumeric(something) Or Information.IsNull(something) Then
        convertToJSON = convertSimpleTypeToJSON(something)
        Exit Function
               
    ElseIf Information.IsArray(something) Then
        convertToJSON = arrayToJSON(something)
        Exit Function
        
    ElseIf Information.TypeName(something) = "Dictionary" Then
        convertToJSON = dictionaryToJSON(something)
        Exit Function
    End If
        Err.Raise "Type not accepted for JSON."
End Function
