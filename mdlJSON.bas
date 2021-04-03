Attribute VB_Name = "mdlJSON"
Option Explicit
Const QUOTES As String = """"
Const ESCAPED_QUOTES As String = "\" & QUOTES

'
' PARSING FUNCTION EXTRACTED FROM:
' https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'

' Parsing variables.
Private p&, token, dic


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
' Parses the JSON.
'''
Function parse(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function

'''
'
'   PRIVATE FUNCTIONS
'
'''

'
' STRINGIFYING
'

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
    Dim dictKeys As Variant: dictKeys = dictionary.keys()
    Dim dictItems As Variant: dictItems = dictionary.items()
    
    For I = 0 To dictionaryCount
        tmpKey = escapeString(dictKeys(I))
        tmpValue = stringify(dictItems(I))
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

'
'
' PARSING
' READ HERE:  https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'

Private Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{":  ParseObj key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function

Private Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function

Private Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function

Private Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function

Private Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

Private Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function