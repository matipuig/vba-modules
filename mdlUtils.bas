Attribute VB_Name = "mdlUtils"
Option Explicit

'''
' Necessary things for the click press.
'''
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

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

'''
' Presses right click.
'''
Public Sub rightClick()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

'''
' Presses left click.
'''
Public Sub leftClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub