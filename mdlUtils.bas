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
'   Gets a random number between both specified numbers.
'''
Public Function getRandomNumber(ByVal min As Single, ByVal max As Single, Optional ByVal isInteger As Boolean = True) As Single
    Randomize
    Dim random As Single: random = Math.Rnd() * (max - min) + min
    getRandomNumber = random
    If isInteger Then
        getRandomNumber = Math.Round(getRandomNumber)
    End If
End Function

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

