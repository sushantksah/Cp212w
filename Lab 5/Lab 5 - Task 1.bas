Attribute VB_Name = "Module1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID:  169060628
' Date: 14/02/2025
' Program title: Flight Finder
' Description:
'===========================================================+

Sub FindFlights()
    ' Define Variables and Array names
    Dim ws As Worksheet
    Dim lastRow As Long, lastCheckRow As Long
    Dim i As Long, j As Long, k As Long
    Dim resultRow As Long
    Dim origins() As String, destinations() As String, flightNumbers() As String
    Dim checkOrigins() As String, checkDestinations() As String
    Dim results() As Variant
    Dim listSize As Long

    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data")
    If ws Is Nothing Then
        MsgBox "Sheet not found! Check the sheet name.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Find last rows needed to iterate to find matches
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCheckRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row

    ' Set up the arrays
    ReDim origins(1 To lastRow - 1)
    ReDim destinations(1 To lastRow - 1)
    ReDim flightNumbers(1 To lastRow - 1)
    ReDim checkOrigins(1 To lastCheckRow - 1)
    ReDim checkDestinations(1 To lastCheckRow - 1)
    
    ' Read flight data, making lowercase to avoid some pairs being missing
    For i = 2 To lastRow
        origins(i - 1) = LCase(Trim(ws.Cells(i, 1).Value))
        destinations(i - 1) = LCase(Trim(ws.Cells(i, 2).Value))
        flightNumbers(i - 1) = ws.Cells(i, 3).Value
    Next i

    ' Checking pairs, making lowercase to avoid some pairs being missing
    For i = 2 To lastCheckRow
        checkOrigins(i - 1) = LCase(Trim(ws.Cells(i, 5).Value))
        checkDestinations(i - 1) = LCase(Trim(ws.Cells(i, 6).Value))
    Next i

    ' Clear previous results
    ws.Range("H2:J" & ws.Rows.Count).ClearContents

    ' Formatting
    With ws.Range("H1:J1")
        .Value = Array("Origin", "Destination", "Flight Number")
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 255)
        .Font.Italic = True
    End With

    ' Storing Results
    listSize = 0
    ReDim results(1 To 3, 1 To 1)

    ' Searching for matches
    ' Don't know why this is skipping over some pairs, tried to debug couldn't understand why
    For i = 1 To UBound(checkOrigins)
        For j = 1 To UBound(checkDestinations)
            For k = 1 To UBound(origins)
                If origins(k) = checkOrigins(i) And destinations(k) = checkDestinations(j) Then
                    listSize = listSize + 1
                    ReDim Preserve results(1 To 3, 1 To listSize)
                    results(1, listSize) = ws.Cells(k + 1, 1).Value
                    results(2, listSize) = ws.Cells(k + 1, 2).Value
                    results(3, listSize) = ws.Cells(k + 1, 3).Value
                End If
            Next k
        Next j
    Next i

    ' If matches found, paste results
    If listSize > 0 Then
        ws.Range("H2").Resize(listSize, 3).Value = Application.Transpose(results)
        Columns("H:J").AutoFit
        MsgBox listSize & " flights found!", vbInformation
    Else
        MsgBox "No matching flights found.", vbExclamation
    End If

    ' Cleanup
    Set ws = Nothing
End Sub

Sub Clear()
    ' Regular Worksheet Clearer
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data")
    On Error GoTo 0
    
    ws.Range("H2:J" & ws.Rows.Count).ClearContents
    
    With ws.Range("H1:J1")
        .Value = Array("Origin", "Destination", "Flight Number")
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
        .Font.Italic = False
    End With
    
    MsgBox "Results cleared successfully!", vbInformation
End Sub
