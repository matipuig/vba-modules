Attribute VB_Name = "mdlJSON"
Option Explicit
Const QUOTES As String = """"
Const ESCAPED_QUOTES As String = "\" & QUOTES

'''
'   Converts the specified received thing in a JSON string.
'''
Public Function stringify(ByRef something As Variant) As String
    Dim someType As Variant: someType = Information.VarType(something)
    
    If Information.IsArray(something) Then
        stringify = arrayToJSON(something)
    
    ElseIf Information.TypeName(something) = "Dictionary" Then
        stringify = dictionaryToJSON(something)
    
    ElseIf Information.TypeName(something) = "Collection" Then
        stringify = collectionToJSON(something)
    
    Else
        stringify = convertSimpleTypeToJSON(something)
    End If
End Function

'''
'   PRIVATE FUNCTIONS
'''

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
        tmpConversion = stringify(someArray(I))
        arrayToJSON = arrayToJSON & tmpConversion
        If I < arraySize Then
            arrayToJSON = arrayToJSON & ","
        End If
    Next I
    arrayToJSON = "[" & arrayToJSON & "]"
End Function

'''
'   Converts a collection to JSON.
'''
Private Function collectionToJSON(ByRef someCollection As Variant) As String
    Dim I As Double
    Dim tmpConversion As String
    Dim collectionSize As Double: collectionSize = someCollection.Count

    For I = 1 To collectionSize
        tmpConversion = stringify(someCollection(I))
        collectionToJSON = collectionToJSON & tmpConversion
        If I < collectionSize Then
            collectionToJSON = collectionToJSON & ","
        End If
    Next I
    collectionToJSON = "[" & collectionToJSON & "]"
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
        tmpValue = stringify(dictionary.items()(I))
        tmpKeyValue = tmpKey & ":" & tmpValue
        If I < dictionaryCount Then
            tmpKeyValue = tmpKeyValue & ","
        End If
        dictionaryToJSON = dictionaryToJSON & tmpKeyValue
    Next I
    dictionaryToJSON = "{" & dictionaryToJSON & "}"
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
        
    ElseIf Information.VarType(something) = vbDate Then
        convertSimpleTypeToJSON = Format(something, "yyyy-mm-dd hh:mm:ss", vbSunday, vbFirstJan1)
        Exit Function
    End If
    
    Err.Raise 1, , "Type " & Information.VarType(something) & " not accepted for JSON."
End Function

'''
'   Escapes the string to return
'''
Private Function escapeString(ByVal text As String) As String
    Dim escapedString As String: escapedString = Replace(text, QUOTES, ESCAPED_QUOTES)
    escapeString = QUOTES & escapedString & QUOTES
End Function
