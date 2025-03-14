Attribute VB_Name = "Task2"
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 06/03/2025
' Program title: Subroutines outside UserForm
' Description: Lab 8
'===========================================================+
' For button to view UserForm for Task 1
Sub viewUserform()
    UserForm1.Show
End Sub

Sub OpenAFile()
    ' Opens a file from the folder this workbook is saved in
    Dim strFileName As String
    
    ' Get the full file location from the user
    strFileName = InputBox("Enter a Full File Location: ")
    

    ' Turn off the default Excel messages
    Application.DisplayAlerts = False

    ' Opening File utilizing error handling
    On Error Resume Next
    ' Check if the file exists before trying to open it
    If Dir(strFileName) = "" Then
        MsgBox "The file '" & strFileName & "' cannot be found or another error occured.", vbCritical, "File Error"
        Exit Sub
    End If
    ' Reset error handling
    On Error GoTo 0
    
    Workbooks.Open Filename:=strFileName

    ' Turn on the alerts again
    Application.DisplayAlerts = True
End Sub
