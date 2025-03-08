VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Receivables 
   Caption         =   "UserForm1"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "Receivables.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Receivables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 07/03/2025
' Program title: private Subroutines Inside UserForm
' Description: Lab 7
'===========================================================+

' Programming Cancel Button
Private Sub Cancel_Click()
    Unload Me
End Sub

' Sub for OK button
Private Sub OK_Click()
    Dim ws As Worksheet
    Dim VTA As String, outputMessage As String, sizeType As String
    Dim lastRow As Long
    Dim cSize As Integer, count As Integer
    Dim numRng As Range, cell As Range
    Dim sum As Double, avg As Double
    
    ' Setting variables needed from the worksheet
    Set ws = ThisWorkbook.Sheets("Receivable")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set numRng = ws.Range("A4:A" & lastRow)
    
    
    ' Get Respose Accoring to what button as clicked (Customer Size)
    If sButton.Value Then
        cSize = 1
    ElseIf mButton.Value Then
        cSize = 2
    ElseIf LButton.Value Then
        cSize = 3
    ' If option is selected
    Else
        MsgBox "Select a cusutomer size!", vbCritical
    End If
    
    ' Get Respose Accoring to what button as clicked (VTA)
    If dButton.Value Then
        VTA = "Days"
    ElseIf aButton.Value Then
        VTA = "Amount"
    Else
        MsgBox "Select an option!", vbCritical
    End If
    
    
    ' Loop through the data to find the average & number of values
    sum = 0
    count = 0
    
    For Each cell In numRng
        'Nested if statements to work with results from UserForm
        If cell.Value = cSize Then
            If VTA = "Days" Then
                ' Add the value of the data that fits the size and
                sum = sum + ws.Cells(cell.Row, 2).Value
            Else
                ' For Amount
                sum = sum + ws.Cells(cell.Row, 3).Value
            End If
            'Increment number of values
            count = count + 1
        End If
    Next cell
    
    ' Find the average
    avg = sum / count
    
    ' Size type for output message
    If cSize = 1 Then
        sizeType = "Small"
    ElseIf cSize = 2 Then
        sizeType = "Medium"
    Else
        sizeType = "Large"
    End If
        
    ' Set up dynamic output message
    outputMessage = "The average of " & VTA & " for " & sizeType & " customers is " & Format(avg, "0.00") & "."
    
    ' Output Message showing results
    MsgBox outputMessage, vbInformation
    
End Sub
