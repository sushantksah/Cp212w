VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Userform1 
   Caption         =   "Select Data"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Test UserForm.frx":0000
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
' Student ID:   169060628
' Date: 03/20/2025
' Program title: Lab 9
' Description: Subroutines inside userform
'===========================================================+

' To Cancel The UserForm
Private Sub cancelButton_Click()
    Unload Me
End Sub

' To Clear the Output
Private Sub clearButton_Click()
    outputBox.Clear
End Sub

' Browse button for the file
Private Sub browseButton_Click()
    Dim fd As FileDialog
    Dim notCancel As Boolean
    
    'Init File Dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Selecting file
    With fd
        .Title = "Select Database File"
        notCancel = .Show
        If notCancel Then
            TextBox1.Value = fd.SelectedItems(1)
        End If
    End With
    
    Set fd = Nothing
    
End Sub

' Opening the file and running the SQL Query
Private Sub runButton_Click()
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim SQL As String, databasePath As String
    Dim countOrders As Integer
    
    ' Get Database Path
    databasePath = TextBox1.Value
    
    '  If no file is selected
    If databasePath = "" Then
        MsgBox "Please select a database file first.", vbExclamation, "No File Selected"
        Exit Sub
    End If
    
    ' SQL Query
    SQL = "SELECT OrderID FROM Orders WHERE CustomerID = 1"
    
    ' Initialize Connection
    With conn
        .ConnectionString = "Data Source =" & databasePath
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    ' Execute Query
    Set rs = conn.Execute(SQL)
    
    ' Populate ListBox
    With rs
        countOrders = 0
        Do While Not rs.EOF
            outputBox.AddItem rs.Fields("OrderID").Value
            countOrders = countOrders + 1
            .MoveNext
        Loop
    End With
    
    ' Output Message
    MsgBox "Total Orders for CustomerID = 1 is: " & countOrders, vbInformation
    
    ' Close Recordset and Connection
    rs.Close
    Set rs = Nothing
    
    conn.Close
    Set conn = Nothing
End Sub
