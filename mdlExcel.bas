Attribute VB_Name = "mdlExcel"
Option Explicit
'''
'   EXCEL METHODS.
'   This module contains methods for using and controlling excel object.
'''

'''
'   Returns the cell address. For example cell(1,1) will return "A1"
'''
Public Function convertToRange(ByRef xlSheet As Object, ByVal rowNumber As Long, ByVal columnNumber As Long) As String
    Dim cell As Object: Set cell = xlSheet.Cells(rowNumber, columnNumber)
    convertToRange = cell.Address(ColumnAbsolute:=False, RowAbsolute:=False)
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
End Function

'''
'   UBICAR FILA CON EL CONTENIDO
'   Revisa la columna en los intervalos especificados y devuelve el n?mero de fila en la cual est? el contenido.
'   Si no lo encuentra, devuelve -1.
'   Puede simplify los contenidos, lo cual quiere decir que search? con trim, y minimizando.
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
    
    searchInColumn = xlSheet.range(foundRange).Row
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


