Attribute VB_Name = "mdlExcel"
Option Explicit
'''
'   EXCEL METHODS.
'   This module contains methods for using and controlling excel object.
'''

'''
'   Opens an Excel App and return its reference.
'''
Public Function getExcelApp(Optional ByVal visible As Boolean = False) As Object
    Static excelApp As Object
    If isApplication(excelApp) = False Then
        Set excelApp = CreateObject("Excel.Application")
        excelApp.visible = visible
    End If
    Set getExcelApp = excelApp
End Function

'''
'   Returns the cell address. For example cell(1,1) will return "A1"
'''
Public Function convertToRange(ByRef xlSheet As Object, ByVal rowNumber As Long, ByVal columnNumber As Long, Optional ByVal absoluteRow As Boolean = False, Optional ByVal absoluteColumn As Boolean = False) As String
    Dim cell As Object: Set cell = xlSheet.Cells(rowNumber, columnNumber)
    convertToRange = cell.Address(ColumnAbsolute:=absoluteColumn, RowAbsolute:=absoluteRow)
    Set cell = Nothing
End Function

'''
'   Executes a search in the specified range and returns the address where it find the result.
'   Returns "" if it cannot find it.
'''
Public Function search(ByRef xlSheet As Object, ByVal range As String, ByVal value As String, Optional ByVal matchCase As Boolean = False) As String
    Dim foundRange As Object
    Set foundRange = xlSheet.range(range).Find(What:=value, LookIn:=-4163, matchCase:=matchCase)
    If foundRange Is Nothing Then
        search = ""
        Exit Function
    End If
    search = foundRange.Address(ColumnAbsolute:=False, RowAbsolute:=False)
    Set foundRange = Nothing
End Function

'''
'   Looks for a specific content in an entire column.
'   Returns the row number if it finds it, or -1.
'''
Public Function searchInColumn(ByRef xlSheet As Object, ByVal searchedValue As String, ByVal column As Long, ByVal startingRow As Long, ByVal endingRow As Long, Optional ByVal matchCase As Boolean = False) As Long
    Dim startingRange As String: startingRange = convertToRange(xlSheet, startingRow, column)
    Dim endingRange As String: endingRange = convertToRange(xlSheet, endingRow, column)
    Dim range As String: range = startingRange & ":" & endingRange
    
    Dim foundRange As String: foundRange = search(xlSheet, range, searchedValue, matchCase)
    If foundRange = "" Then
        searchInColumn = -1
        Exit Function
    End If
    
    searchInColumn = xlSheet.range(foundRange).row
End Function

'''
'   Looks for a specific content in an entire row.
'   Returns the column number if it finds it, or -1.
'''
Public Function searchInRow(ByRef xlSheet As Object, ByVal searchedValue As String, ByVal row As Long, ByVal startingCol As Long, ByVal endingCol As Long, Optional ByVal matchCase As Boolean = False) As Long
    Dim startingRange As String: startingRange = convertToRange(xlSheet, row, startingCol)
    Dim endingRange As String: endingRange = convertToRange(xlSheet, row, endingCol)
    Dim range As String: range = startingRange & ":" & endingRange
    
    Dim foundRange As String: foundRange = search(xlSheet, range, searchedValue, matchCase)
    If foundRange = "" Then
        searchInRow = -1
        Exit Function
    End If
    
    searchInRow = xlSheet.range(foundRange).column
End Function


'''
'   Gets the first empty row in the specified column between the specified rows. If you choose 'trimContent', " " will be interpreted as empty.
'   If it cannot find one, returns -1.
'''
Public Function getNextEmptyRow(ByRef xlSheet As Object, ByVal column As Long, ByVal startingRow As Long, ByVal endingRow As Long, Optional ByVal trimContent As Boolean = True) As Long
    If endingRow < startingRow Then
        Err.Raise 1, , "Ending row should be higher than starting row."
    End If
    
    getNextEmptyRow = -1
    Dim actualRow As Long: actualRow = startingRow
    Dim tmpContent As String

    Do While actualRow <= endingRow
        tmpContent = xlSheet.Cells(actualRow, column).value
        If trimContent = True Then
            tmpContent = Strings.Trim(tmpContent)
        End If
        If tmpContent = "" Then
            getNextEmptyRow = actualRow
            Exit Function
        End If
        actualRow = actualRow + 1
    Loop
End Function

'''
'   Gets the first empty column in the specified column between the specified rows. If you choose 'trimContent', " " will be interpreted as empty.
'   If it cannot find one, returns -1.
'''
Public Function getNextEmptyColumn(ByRef xlSheet As Object, ByVal row As Long, ByVal startingCol As Long, ByVal endingCol As Long, Optional ByVal trimContent As Boolean = True) As Long
    If endingCol < startingCol Then
        Err.Raise 1, , "Ending col should be higher than starting col."
    End If
    
    getNextEmptyColumn = -1
    Dim actualCol As Long: actualCol = startingCol
    Dim tmpContent As String

    Do While actualCol <= endingCol
        tmpContent = xlSheet.Cells(row, actualCol).value
        If trimContent = True Then
            tmpContent = Strings.Trim(tmpContent)
        End If
        If tmpContent = "" Then
            getNextEmptyColumn = actualCol
            Exit Function
        End If
        actualCol = actualCol + 1
    Loop
End Function

'''
'   Creates a collection from the row.
'   Example: ["cell1", "cell2", "cell3", ...]
'''
Public Function rowToCollection(ByRef xlSheet As Object, ByVal row As Long, ByVal startingCol As Long, ByVal endingCol As Long, Optional ByVal toUpperCase As Boolean = False) As Object
    If endingCol < startingCol Then
        Err.Raise 1, , "Ending col should be higher than starting col."
    End If
    
    Set rowToCollection = New Collection
    Dim I As Long
    Dim tmpKey As String
    For I = startingCol To endingCol
        tmpKey = Conversion.CStr(xlSheet.Cells(row, I).value)
        If toUpperCase Then
            tmpKey = Strings.UCase(tmpKey)
        End If
        rowToCollection.Add tmpKey
    Next I
End Function

'''
'   Creates a dictionary containing each column text in the row and the column number.
'   Example: Dictionary: {value1: 1, value2: 2, value3: 3, value4: 4, etc.}
'   Then you can loop through: For I = 0 to Dictionary.count: Msgbox Dictonary.Keys()(I) & " = " & Dictionary.Items()(I)
'''
Public Function rowToIndexDictionary(ByRef xlSheet As Object, ByVal row As Long, ByVal startingCol As Long, ByVal endingCol As Long, Optional ByVal toUpperCase As Boolean = False) As Object
    If endingCol < startingCol Then
        Err.Raise 1, , "Ending col should be higher than starting col."
    End If
    
    Dim dictionary As Object: Set dictionary = CreateObject("Scripting.Dictionary")
    Dim I As Long
    Dim tmpKey As String
    For I = startingCol To endingCol
        tmpKey = Conversion.CStr(xlSheet.Cells(row, I).value)
        If Strings.Trim(tmpKey) = "" Then
            GoTo nextCol
        End If
        If toUpperCase Then
            tmpKey = Strings.UCase(tmpKey)
        End If
        dictionary(tmpKey) = I
nextCol:
    Next I
    Set rowToDictionary = dictionary
End Function

'''
'   Creates a collection of dictionaries of the specified sheet area, using the first row as headers and creating a new collection dictionary item per row.
'   Headers are taken from the first row.
'   Example: [Dictionary (row 2): {header1: "Value 1", header2: "value 2", header3: "Value 3"}, Dictionary (row 3), Dictionary (row 4)...]
'''
Public Function sheetToDictionaryCollection(ByRef xlSheet As Object, ByVal startingRow As Long, ByVal endingRow As Long, ByVal startingCol As Long, ByVal endingCol As Long, Optional ByVal headersToUpperCase As Boolean = False) As Collection
    If endingRow < startingRow Then
        Err.Raise 1, , "Ending row should be higher than starting row."
    End If
    If endingCol < startingCol Then
        Err.Raise 1, , "Ending col should be higher than starting col."
    End If
      
    Set sheetToDictionaryCollection = New Collection
      
    Dim firstContentRow As Double: firstContentRow = startingRow + 1
    If firstContentRow > endingRow Then
        Exit Function
    End If
    
    Dim headers As Collection: Set headers = rowToCollection(xlSheet, startingRow, startingCol, endingCol, headersToUpperCase)
    Dim I As Double
    Dim J As Double
    Dim header As String
    Dim rowDictionary As Object
    For I = firstContentRow To endingRow
        Set rowDictionary = CreateObject("Scripting.Dictionary")
        
        For J = startingCol To endingCol
            header = headers(J - startingCol + 1)
            rowDictionary.Add header, xlSheet.Cells(I, J).value
            DoEvents
        Next J
        
        sheetToDictionaryCollection.Add rowDictionary
        DoEvents
    Next I
    
    Set rowDictionary = Nothing
End Function


'''
'   Closes the excel app.
'''
Public Function closeApp() As Boolean
    Dim excelApp As Object: Set excelApp = getExcelApp
    If isApplication(excelApp) = False Then
        closeApp = True
        Exit Function
    End If
        
    Dim wb
    For Each wb In excelApp.Workbooks
        wb.Close SaveChanges:=False
    Next wb
    excelApp.Quit
    Set excelApp = Nothing
    closeApp = True
End Function


'
'
'   PRIVATE METHODS
'
'

'''
'   Checks if is open.
'''
Private Function isApplication(ByRef object) As Boolean
    isApplication = False
    If Information.TypeName(object) = "Application" Then
        isApplication = True
    End If
End Function

