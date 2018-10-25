Attribute VB_Name = "mainModule"
Option Explicit

Public Sub main()
    Dim columns(1 To 2) As Integer
    Dim startTime As Double
    Dim stopWatch As Double
    Dim result As New Collection
    Dim startRow As Long
    Dim str As String
    Dim i As Long
    
    Debug.Print ""
    Debug.Print "==========================================================================="
    Debug.Print "Comparison started.  Loading Sheet1 & Sheet2..."
    startTime = Timer()
    
    columns(1) = 3
    columns(2) = 4
    startRow = 88
    Dim sheet1 As classCompareSheet
    Set sheet1 = factory.createCompareSheet(ThisWorkbook.Sheets(1), startRow, columns)
    
    columns(1) = 1
    columns(2) = 2
    startRow = 1
    Dim sheet2 As classCompareSheet
    Set sheet2 = factory.createCompareSheet(ThisWorkbook.Sheets(2), startRow, columns)
    
    Debug.Print "Sheets loaded.  Time Elapsed: " & Timer() - startTime & " seconds"
    Debug.Print ""
    
    stopWatch = Timer()
    Debug.Print "Comparing Sheet1 to Sheet2"
    Set result = sheet1.checkAgainst(sheet2, 0)
    Debug.Print "On Sheet1 but NOT on Sheet2: " & result.Count & " entries"
    Debug.Print "It took " & Timer() - stopWatch & " seconds"
    
    str = ""
    For i = 1 To result.Count
        If i = 1 Then
            str = result(i)
        Else
            str = str & "," & result(i)
        End If
    Next
    Debug.Print "They are found on the following lines: " & str
    Debug.Print ""
    
    stopWatch = Timer()
    Debug.Print "Comparing Sheet2 to Sheet1"
    Set result = sheet1.checkAgainst(sheet2, 1)
    Debug.Print "On Sheet1 AND on Sheet2: " & result.Count & " entries"
    Debug.Print "It took " & Timer() - stopWatch & " seconds"
        str = ""
    For i = 1 To result.Count
        If i = 1 Then
            str = result(i)
        Else
            str = str & "," & result(i)
        End If
    Next
    Debug.Print "They are found on the following lines: " & str
    Debug.Print ""
    
    Debug.Print "Comparison Complete!"
    Debug.Print "Total Time Elapsed: " & Timer() - startTime & " seconds"
End Sub







