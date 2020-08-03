Attribute VB_Name = "mdlInternet"
Option Explicit
'''
'   INTERNET METHODS
'   This module contains all the methods we need for internet.
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
'   Executes an HTTP Request.
'   Nonce is in order to prevent cache.
'   Headers should be a dictionary.
'''
Public Function request(ByVal url As String, ByVal method As String, ByVal body As String, Optional ByVal addNonce As Boolean = False, Optional ByRef headers As Object = Nothing) As String
    Dim xmlHTTP As Object: Set xmlHTTP = getXMLHTTP
    Dim I As Integer

    'Prevent cache with nonce.
    If addNonce = True Then
        Randomize
        Dim nonce As Single: nonce = Rnd(50000) * 40000
        If InStr(url, "?") < 1 Then url = url & "?"
        url = url & "&nonce" & nonce & "=" & nonce
    End If
    
    method = Strings.UCase(Strings.Trim(method))
    xmlHTTP.Open method, url, True
    xmlHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    xmlHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"

    If Not headers Is Nothing Then
        For I = 0 To headers.Count - 1
            xmlHTTP.setRequestHeader headers.keys()(I), headers.items()(I)
        Next I
    End If

    xmlHTTP.send (body)
    DoEvents
    Do While xmlHTTP.readystate <> 4
        DoEvents
    Loop
    request = xmlHTTP.responsetext
    Set xmlHTTP = Nothing
End Function


'''
'   URL encodes a dictionary.
'   Returns the url encoded string.
'''
Public Function urlEncodeDictionary(ByRef dictionary as Object) as String
    Dim I as Integer
    Dim tmpValue As String 
    For I = 0 To dictionary.Count - 1
        tmpValue = urlEncodeKeyAndValue(dictionary.keys()(I), dictionary.items()(I))
        urlEncodeDictionary = urlEncodeDictionary & tmpValue
    Next I
End Function

'''
'   Encodes a key value pair: key = value to "&key=value".
'''
Public Function urlEncodeKeyAndValue(ByVal key As String, ByVal value As String) As String
    key = urlEncode(key)
    value = urlEncode(value)
    urlEncodeKeyAndValue = "&" & key & "=" & value
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
'   Gets the HTML file and its reference.
'''
Private Function getHTML() As Object
    Static HTMLFile As Object
    If HTMLFile Is Nothing Then
        Set HTMLFile = CreateObject("htmlfile")
    End If
    Set getHTML = HTMLFile
End Function


