Attribute VB_Name = "mdlUtils"
Option Explicit
'''
'   UTILS.
'   This module contains some methods are often required.
'''

'''
'   Throws an error with the specified code and message.
'''
Public Sub throwError(ByVal code As Integer, ByVal message As String)
    Err.Raise code, , message
End Sub

'''
'   Wait the specified seconds.
'''
Public Function waitSeconds(ByVal seconds As Long) As Boolean
    Dim stopWhen As Long: stopWhen = Timer + seconds
    Do While Timer <= stopWhen
        DoEvents
    Loop
    waitSeconds = True
End Function
