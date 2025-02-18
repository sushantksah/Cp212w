Attribute VB_Name = "Task2"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 2025/01/24
' Program title: Lab 1 Module 2
' Description: BMI Calculator using user input.
'===========================================================+

Sub BMI()
Dim HeightInput As String
Dim userHeight As Single
Dim WeightInput As String
Dim userWeight As Single
Dim BMI As Single




HeightInput = InputBox("Enter height in metres(m): ")
WeightInput = InputBox("Enter weight in kilograms(kg): ")

userHeight = CSng(HeightInput)
userWeight = CSng(WeightInput)


BMI = ((userWeight) / (userHeight ^ 2))
BMI = Round(BMI, 2)

MsgBox "Body Mass Index Calculation " & vbNewLine & vbCrLf & _
"Height: " & userHeight & vbCrLf & _
"Weight " & userWeight & vbCrLf & _
"BMI: " & BMI


End Sub

