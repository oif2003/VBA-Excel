Attribute VB_Name = "factory"
Option Explicit

'This is where we keep factory methods for Class Modules
'

'Factory method for classCompareSheet, essentially a wrapper for initializeClass() inside classCompareSheet
Public Function createCompareSheet(ws As Worksheet, firstRow As Long, _
                                    columns() As Integer, Optional pipeStr As String) As classCompareSheet
    Dim sheet As New classCompareSheet
    Call sheet.initializeClass(ws, firstRow, columns, pipeStr)
    Set createCompareSheet = sheet
End Function

