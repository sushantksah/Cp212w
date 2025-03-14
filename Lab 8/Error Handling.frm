VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Error Handling.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
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
' Description: Lab 8
'===========================================================+

' Sub for cancel button
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub txtLastName_Change()
    Dim name As String
    name = txtLastName.Text

    If Not (UCase(Right(name, 1)) >= "A" And UCase(Right(name, 1)) <= "Z") Then
        MsgBox "Please enter alphabetical characters only.", vbExclamation
        txtLastName.Text = Left(name, Len(name) - 1)
        txtLastName.SetFocus
        Exit Sub
    End If
End Sub
