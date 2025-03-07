VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TriangularDist 
   Caption         =   "UserForm1"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "TriangularDist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TriangularDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 06/03/2025
' Program title: private Subroutines Inside UserForm
' Description: Assignment 3
'===========================================================+

'To cancel the userform
Private Sub Cancel_Click()
    Unload Me
End Sub

' Submit button
' This sub calls genRand feeding in the user inputted parameters drawn from the UserForm
' Tests to see if the input is useable or not.
Private Sub Submit_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Results")

    ' Store parameter labels in A1:A4
    ws.Range("A1").Value = "Function Parameters:"
    ws.Range("A2").Value = "Minimum (a):"
    ws.Range("A3").Value = "Most Likely (b):"
    ws.Range("A4").Value = "Maximum (c):"
    ws.Range("A5").Value = "Total # of Values: "
    
    ' Formatting Column A
    ws.Columns("A").AutoFit
    With Range("A1").Font
        .Bold = True
        .Size = 14
    End With
    
    With Range("A2:A5").Font
        .Bold = True
        .Color = vbBlue
    End With
    
    ' Input Validation
    If Not Validate(aTextbox.Value) Then Exit Sub
    If Not Validate(bTextbox.Value) Then Exit Sub
    If Not Validate(cTextbox.Value) Then Exit Sub

    ' Printing Values entered into the userform by the user
    ws.Range("B2").Value = aTextbox.Value
    ws.Range("B3").Value = bTextbox.Value
    ws.Range("B4").Value = cTextbox.Value
    ws.Range("B5").Value = mTextbox.Value
    
    

    ' Calling random triangular dist number generator
    genRand
    
    ' Hiding userform after use
    Me.Hide
End Sub

' Input Validation / Error Handling,
' Clearing data as well, to not show inaccurate data (would show and keep updating #'s becausue page would..
' ... refresh and use old parameters, so decided to clear with error as to not confuse)
Function Validate(ByVal inputVal As Double) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Results")
    
    ' See if the value is a Number
    If Not IsNumeric(inputVal) Then
        MsgBox "Input must be a numeric value!", vbCritical
        ws.Range("A7:A1000").ClearContents
        Validate = False
    
    ' Number must be greater than 0
    ElseIf inputVal < 0 Then
        MsgBox "Input must be greater than 0!", vbCritical
        ws.Range("A7:A1000").ClearContents
        Validate = False
    
    ' Number must be less than 100
    ElseIf inputVal > 100 Then
        MsgBox "Input must be less than or equal to 100!", vbCritical
        ws.Range("A7:A1000").ClearContents
        Validate = False
        
    ' If everything is ok, then continue
    Else
        Validate = True
    End If
End Function

Private Sub UserForm_Click()

End Sub
