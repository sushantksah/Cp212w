Attribute VB_Name = "Task1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 2025/01/24
' Program title: Lab 1 Module 1
' Description: Formatting exam data using msgbox and worksheet functions.
'===========================================================+

Sub ExamScores()
Dim rng As Range
Dim maxVal As Double
Dim avg As Double
Dim minVal As Double
Dim stDev As Double



Set rng = Worksheets("ExamScores").Range("A1:A100")

maxVal = Application.WorksheetFunction.Max(rng)
minVal = Application.WorksheetFunction.Min(rng)
stDev = Application.WorksheetFunction.stDev(rng)
avg = Application.WorksheetFunction.Average(rng)

stDev = Round(stDev, 2)
avg = Round(avg, 2)


MsgBox "Here are summary measures for the scores: " & vbCrLf & _
"Average: " & avg & vbCrLf & _
"Stdev: " & stDev & vbCrLf & _
"Min: " & minVal & vbCrLf & _
"Max: " & maxVal, vbInformation

End Sub
