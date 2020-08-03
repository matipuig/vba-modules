Attribute VB_Name = "mdlAdobeAcrobat"
Option Explicit

'''
'   ADOBE ACROBAT METHODS
'   This module lets you interact with Adobe Acrobat.
'   Remember that to use this module you should have Adobe Acrobat, and even you should need to import the reference.
'''

'''
'   Opens the app and returns the object.
'''
Private Function getAcroApp() As Object
    Static acroApp As Object
    If acroApp Is Nothing Then
        Set acroApp = CreateObject("AcroExch.App")
    End If
    Set getAcroApp = acroApp
End Function

''
'   Reads the file content.
''
Public Function getContent(ByVal filePath As String) As String
    Dim acroApp As Object: Set acroApp = getAcroApp
    
    'Note: A Reference to the Adobe Library must be set in Tools|References!
    Dim AcroAVDoc As CAcroAVDoc, AcroPDDoc As CAcroPDDoc
    Dim AcroHiliteList As CAcroHiliteList, AcroTextSelect As CAcroPDTextSelect
    Dim PageNumber, PageContent, Content, I, J
    
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    If AcroAVDoc.Open(filePath, vbNull) <> True Then Exit Function
    ' The following While-Wend loop shouldn't be necessary but timing issues may occur.
    
    While AcroAVDoc Is Nothing
        DoEvents
        Set AcroAVDoc = acroApp.GetActiveDoc
    Wend
    
    Set AcroPDDoc = AcroAVDoc.GetPDDoc
    For I = 0 To AcroPDDoc.GetNumPages - 1
        DoEvents
        Set PageNumber = AcroPDDoc.AcquirePage(I)
        Set PageContent = CreateObject("AcroExch.HiliteList")
        If PageContent.Add(0, 9000) <> True Then Exit Function
        Set AcroTextSelect = PageNumber.CreatePageHilite(PageContent)
        
        ' The next line is needed to avoid errors with protected PDFs that can't be read.
        On Error Resume Next
        For J = 0 To AcroTextSelect.GetNumText - 1
            DoEvents
            Content = Content & AcroTextSelect.GetText(J)
        Next J
    Next I
    
    getContent = Content
    AcroAVDoc.Close True
    Set AcroAVDoc = Nothing
    Set acroApp = Nothing
End Function


'''
'   Close the Adobe Acrobat App.
'''
Public Function closeApp() As Boolean
    Dim acroApp As Object: Set acroApp = getAcroApp
    acroApp.Exit
    Set acroApp = Nothing
    closeApp = True
End Function




