Attribute VB_Name = "Task2"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 10/02/2025
' Program title: Assignment 2 - Part B
' Description: Generate Report based on user inputted number.
'===========================================================+

Sub GenerateReport()
    ' Variable Definition
    Dim wsData As Worksheet, wsReport As Worksheet
    Dim dict As Object
    Dim lastRow As Long, i As Long
    Dim totalAmount As Double
    Dim userInput As Variant
    Dim rowIndex As Integer
    Dim rng As Range
    Dim key As Variant
    Dim custID As Variant, amount As Variant
    
    
    ' Set Data Worksheet as reference for sub
    Set wsData = ThisWorkbook.Sheets("Data1")

    ' Loop until user provides a valid amount or cancels
    Do
        userInput = Application.InputBox("Enter a total amount (i.e., 3000):", Type:=1)
        If userInput = False Then Exit Sub
    Loop Until IsNumeric(userInput) And userInput > 0

    ' Make input a Double
    totalAmount = CDbl(userInput)

    ' Create dictionary to store total amounts per CustomerID
    Set dict = CreateObject("Scripting.Dictionary")

    ' Find last row of Data
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' Iterate through Data sheet and sum orders by Customer ID
    For i = 2 To lastRow
        custID = wsData.Cells(i, 2).Value
        
        If Not IsEmpty(custID) And Not IsError(custID) Then
            custID = CStr(custID)
            If IsDate(custID) Then
                custID = "'" & custID
            End If
            amount = wsData.Cells(i, 3).Value
            If IsNumeric(amount) Then
                amount = CDbl(amount)
                If dict.exists(custID) Then
                    dict(custID) = dict(custID) + amount
                Else
                    dict.Add custID, amount
                End If
            End If
        End If
    Next i

    ' Delete existing "Report" worksheet if it exists (Error Handling)
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsReport = ThisWorkbook.Sheets("Report")
    If Not wsReport Is Nothing Then wsReport.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create new "Report" worksheet
    Set wsReport = ThisWorkbook.Sheets.Add
    wsReport.Name = "Report"

    ' Add headers with proper formatting
    With wsReport
        .Cells(1, 1).Value = "Customers who spent more than $" & Format(totalAmount, "#,##0.00")
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 12
        .Cells(3, 1).Value = "Customer ID"
        .Cells(3, 2).Value = "Total Amount Spent"
        .Range("A3:B3").Font.Bold = True
        .Range("A3:B3").Font.Size = 11
    End With

    ' Clear any existing formatting in the columns before entering values
    wsReport.Columns("A:B").ClearFormats

    ' Initialize rowIndex before loop, starting from row four as row 2 and 3
    ' are already set as blank space and headings
    rowIndex = 4

    ' Populate report with customers who spent more than the user-inputted value
    ' of totalAmount
    For Each key In dict.Keys
        If dict(key) > totalAmount Then
            wsReport.Cells(rowIndex, 1).Value = key
            wsReport.Cells(rowIndex, 2).Value = dict(key)
            wsReport.Cells(rowIndex, 2).NumberFormat = "$#,##0.00"
            wsReport.Columns("A").ColumnWidth = 15
            wsReport.Columns("B").AutoFit
            rowIndex = rowIndex + 1
        End If
    Next key

    ' Sort the report in ascending order by total amount if there are entries
    If rowIndex > 4 Then
        Set rng = wsReport.Range("A3:B" & rowIndex - 1)
        rng.Sort Key1:=wsReport.Range("B4"), Order1:=xlAscending, Header:=xlYes
    End If

    ' Cleanup
    Set dict = Nothing
    MsgBox "Totals Report generated successfully!", vbInformation
End Sub
