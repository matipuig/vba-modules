Attribute VB_Name = "mdlInternet"
Option Explicit
'''
'   INTERNET METHODS
'   This module contains all the methods we need for internet.
'''

'''
'   Executes an HTTP Request.
'   Nonce is in order to prevent cache.
'   Headers should be a dictionary.
'''
Public Function request(ByVal url As String, ByVal method As String, ByVal body As String, Optional ByVal addNonce As Boolean = False, Optional ByRef headers As Object = Nothing, Optional ByVal isJSON As Boolean = False) As String
    Dim xmlHTTP As Object: Set xmlHTTP = getXMLHTTP
    Dim I As Integer

    'Prevent cache with nonce.
    If addNonce = True Then
        Randomize
        Dim nonce As Single: nonce = Rnd(50000) * 40000
        If Strings.InStr(url, "?") < 1 Then url = url & "?"
        url = url & "&nonce" & nonce & "=" & nonce
    End If
    
    method = Strings.UCase(Strings.Trim(method))
    xmlHTTP.Open method, url, True
    xmlHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    
    If Not headers Is Nothing Then
        For I = 0 To headers.Count - 1
            xmlHTTP.setRequestHeader headers.keys()(I), headers.items()(I)
        Next I
    End If
    
    If isJSON = True Then
        xmlHTTP.setRequestHeader "Content-Type", "application/json"
    Else
        xmlHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    End If

    xmlHTTP.send (body)
    DoEvents
    Do While xmlHTTP.readyState <> 4
        DoEvents
    Loop
    request = xmlHTTP.responseText
    Set xmlHTTP = Nothing
End Function

'''
'   URL encodes a dictionary.
'   Returns the url encoded string.
'''
Public Function urlEncodeDictionary(ByRef dictionary As Object) As String
    Dim I As Integer
    Dim tmpKey As String
    Dim tmpValue As String
    Dim tmpKeyValue As String
    For I = 0 To dictionary.Count - 1
        tmpKey = urlEncode(dictionary.keys()(I))
        tmpValue = urlEncode(dictionary.items()(I))
        tmpKeyValue = "&" & tmpKey & "=" & tmpValue
        urlEncodeDictionary = urlEncodeDictionary & tmpKeyValue
    Next I
End Function

'''
'   URL encodes some text.
'''
Public Function urlEncode(ByVal text As String) As String
    Static urlEncodeCreated As Boolean
    Dim HTMLFile As Object: Set HTMLFile = getHTML
    If urlEncodeCreated = False Then
        HTMLFile.parentWindow.execScript "function urlEncode(text) {return encodeURIComponent(text);}", "jscript"
        urlEncodeCreated = True
    End If
    urlEncode = HTMLFile.parentWindow.urlEncode(text)
    Set HTMLFile = Nothing
End Function

'''
'
'   PRIVATE METHODS.
'
'''

'''
'   Creates and return an XMLHTTP Object for requests.
'''
Private Function getXMLHTTP() As Object
    Static xmlHTTP As Object
    If xmlHTTP Is Nothing Then
        Set xmlHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    End If
    Set getXMLHTTP = xmlHTTP
End Function

'''
'   Gets the HTML file and its reference.
'''
Private Function getHTML() As Object
    Static HTMLFile As Object
    If HTMLFile Is Nothing Then
        Set HTMLFile = CreateObject("htmlfile")
    End If
    Set getHTML = HTMLFile
End Function



