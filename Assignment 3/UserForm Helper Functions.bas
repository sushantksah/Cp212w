Attribute VB_Name = "Assignment3"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 06/03/2025
' Program title: Subroutines outside UserForm
' Description: Assignment 3
'===========================================================+

' For button to view UserForm
Sub viewUserform()
    TriangularDist.Show
End Sub

'Trinagulaar Distribution Function (Task 1)
Function Triangular(a As Double, b As Double, c As Double) As Double
Dim d As Double, u As Double

' Randomization / Recalculation, ensuring the function will recalculate when prompted
Randomize
Application.Volatile

' Equation for D
d = (b - a) / (c - a)

' Setting u to a random # 0-1
u = Rnd

' Calculating the random number based on previous equation
If u <= d Then
    Triangular = a + (c - a) * Sqr(d * u)

ElseIf u > d Then
    Triangular = a + (c - a) * (1 - Sqr((1 - d) * (1 - u)))
    
End If
End Function

' Sub to generate random values other than min, max, and most likely
Sub genRand()
    Dim ws As Worksheet
    Dim minVal As Double, mostLikely As Double, maxVal As Double
    Dim totalValues As Integer, i As Integer, rowStart As Integer
    
    ' Setting workbook to correct one
    Set ws = ThisWorkbook.Sheets("Results")
    
    ' Clear ws from prev uses
    ws.Range("A7:A1000").ClearContents
    
    rowStart = 7
    
    ' Adding in the values of min/max
    minVal = ws.Range("B2").Value
    maxVal = ws.Range("B4").Value
    totalValues = ws.Range("B5").Value
    
    
    ' changing calculation mode to manual so f9 can be pressed to print new numbers
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Fill the cells with the random numbers
    ' Didn't know what was meant by "you will write your function into the introduction sheet not using the worksheet function technique"
    For i = 0 To totalValues - 1
        ws.Cells(rowStart + i, 1).Formula = "=Triangular(B2, B3, B4)"
    Next i
    
    ' Putting everything back to Normal
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
