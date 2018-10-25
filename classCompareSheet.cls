VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classCompareSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pColl As New Collection
Private pSheet As Worksheet
Private pLastRow As Long
Private pName As String
Private pFirstRow As Long
Private pPipeString As String
Private pCollArr() As String

Public Sub initializeClass(ws As Worksheet, firstRow As Long, columns() As Integer, Optional pipeStr As String)
    If pipeStr = Null Then
        pPipeString = "|"
    Else
        pPipeString = pipeStr
    End If
    
    Set pSheet = ws
    pName = pSheet.name
    pLastRow = findLastRow()
    pFirstRow = firstRow
    Call createCollection(columns)
End Sub

Public Property Let name(value As String)
    pName = value
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Get lastRow() As Long
    lastRow = pLastRow
End Property

Public Function entryExist(strkey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = pColl(strkey)
    entryExist = (Err.Number = 0)
    Err.Clear
End Function

Public Function checkAgainst(oCompareSheet As classCompareSheet, flag As Integer) As Collection
    Dim result As New Collection
    'Dim element As Variant
    Dim i As Long
    
    If flag = 1 Then
        For i = LBound(pCollArr) To UBound(pCollArr)
            If oCompareSheet.entryExist(pCollArr(i)) Then
                result.Add (i + pFirstRow - 1)
            End If
        Next i
        
    ElseIf flag = 0 Then
        For i = LBound(pCollArr) To UBound(pCollArr)
            If Not oCompareSheet.entryExist(pCollArr(i)) Then
                result.Add (i + pFirstRow - 1)
            End If
        Next i
    
    Else
        Debug.Print "Flag parameter should be 0 (not in), or 1(is in)"
        Exit Function
    End If
    
    Set checkAgainst = result
End Function

Private Sub createCollection(columns() As Integer)
    Dim uniqueID As String
    Dim i As Long
    Dim j As Integer
    Dim k As Long
    Dim sheet() As Variant
    
    sheet = pSheet.Range(pSheet.cells(pFirstRow, columns(LBound(columns))), _
                         pSheet.cells(pLastRow, columns(UBound(columns))))
    ReDim pCollArr(1 To pLastRow - pFirstRow + 1)
    k = 1
    For i = 1 To pLastRow - pFirstRow + 1
        uniqueID = ""
        For j = LBound(columns) To UBound(columns)
            uniqueID = uniqueID & sheet(i, j) & pPipeString
        Next j

        Do While entryExist(uniqueID)
            uniqueID = uniqueID & "."
        Loop
        pCollArr(i) = uniqueID
        k = k + 1
        'use VBA.Array so "Option Base 1" will not effect us
        Call pColl.Add(Null, key:=uniqueID)
    Next i

End Sub

'https://www.excelcampus.com/vba/find-last-row-column-cell/
Private Function findLastRow() As Long
    'first check to see if sheet is empty, otherwise cells.Find will return nothing
    If Application.WorksheetFunction.CountA(pSheet.cells) > 0 Then
        findLastRow = pSheet.cells.Find(What:="*", After:=Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                                        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
    Else
        findLastRow = 1
    End If
End Function
