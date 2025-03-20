VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Task2 
   Caption         =   "Type in a letter"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "Select Case.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Task2"
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
' Description: Assignment 4
'===========================================================+

'To cancel the userform
Private Sub cancelButton_Click()
    Unload Me
End Sub

' Submit button, changing cases
Private Sub submitButton_Click()
    Dim rng As Range, cell As Range
    Dim choice As String
    Dim words As Variant

' ==== Error Handling ======================================+
   ' Ensure a range is selected prior to running the rest
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cells to change!", vbCritical
        Exit Sub
    End If

    ' Setting range to be looped through as the selection
    Set rng = Selection

    ' Check if selection has text using CountA that will see if there are letters or numbers in the cells
    ' The cells will be checked later to check if it isn't empty/isn't numbers
    If WorksheetFunction.CountA(rng) = 0 Then
        MsgBox "Selected range contains no text!", vbCritical
        Exit Sub
    End If

    ' Get input from textbox, making uppercase so the input isn't case sensative
    ' Make sure input is one of the choices provided, if not send error message
    choice = UCase(inputBox.Value)
    If InStr("LUSTC", choice) = 0 Then
        MsgBox "Input Invalid! Please enter either L, U, S, T, or C.", vbCritical
        Exit Sub
    End If
'===========================================================+

    ' Loop to assess each case
    For Each cell In rng
        ' Base case ensuring that the cells have text
        If Not IsEmpty(cell.Value) And Not IsNumeric(cell.Value) Then
            Select Case choice
            
                ' Lowercase Case
                Case "L"
                    cell.Value = LCase(cell.Value)

                'Uppercase Case
                Case "U"
                    cell.Value = UCase(cell.Value)
            
                ' Sentance Case
                Case "S"
                    cell.Value = UCase(Left(cell.Value, 1)) & LCase(Mid(cell.Value, 2))
                    
                ' Title Case
                Case "T"
                    cell.Value = Application.WorksheetFunction.Proper(LCase(cell.Value))

                ' Caps Small Case
                Case "C"
                ' Couldn't get it to work properly on time
                ' But would make the input uppercase, start at index 2 in the string after converting it to
                ' chars and make the font lower using a loop
                
            End Select
        End If
    Next cell

    MsgBox "Case change applied successfully!", vbInformation

    ' Close userform
    Unload Me
End Sub
