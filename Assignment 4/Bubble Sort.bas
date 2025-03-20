Attribute VB_Name = "Task1"
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 19/03/2025
' Program title: Subroutines for Assignment 4 (Outside Userform)
' Description: Assignment 4
'===========================================================+

' Task 1 - Bubble Sort in VBA
Sub vbaBubbleSort()
    Dim ws As Worksheet
    Dim startRow As Integer, lastRow As Integer
    Dim temp As Variant
    Dim i As Integer, j As Integer
    Dim data() As Variant

    ' Set ws
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Start at a2 as a1 is header
    startRow = 2
    
    ' Find lastRow
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Resize array to accomodate column values dynamically
    ReDim data(startRow To lastRow)
    For i = startRow To lastRow
        data(i) = ws.Cells(i, 1).Value
    Next i

    ' Bubble Sort Implementation, iterating through the array comparing adjacent numbers
    For i = startRow To lastRow - 1
        For j = startRow To lastRow - (i - startRow) - 1
            If data(j) > data(j + 1) Then
                ' Swapping positions if num is greater
                temp = data(j)
                data(j) = data(j + 1)
                data(j + 1) = temp
            End If
        Next j
    Next i

    ' Populate from b2 onwards with the sorted results
    For i = startRow To lastRow
        ws.Cells(i, 2).Value = data(i)
    Next i

End Sub

' Task 2 - Change text case based on user input/Case function
Sub viewUserform()
    Task2.Show
End Sub

