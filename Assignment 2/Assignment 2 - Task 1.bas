Attribute VB_Name = "Task1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 10/02/2025
' Program title: Assignment 2 - Part A
' Description: Creates worksheets based on product categories
'===========================================================+
Sub CreateCategorySheets()
    Dim wsData As Worksheet, ws As Worksheet
    Dim lastRow As Long, rowCount As Integer
    Dim cell As Range
    Dim cat As String
    
    ' Set the Data worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("Data")
    On Error GoTo 0
    
    ' Find the last row in column B
    lastRow = wsData.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' Loop through column B to create different category sheets
    For Each cell In wsData.Range("B4:B" & lastRow)
        cat = CleanSheetName(cell.Value)
        
        ' Error Handling
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(cat)
        On Error GoTo 0
        
        ' If sheet doesn't exist, create it
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add
            On Error Resume Next
            ws.Name = cat
            If Err.Number <> 0 Then
                MsgBox "Error creating sheet: " & cat, vbCritical
                Err.Clear
            End If
            On Error GoTo 0
        End If
        
        ' Add headers if empty
        If ws.Range("A1").Value = "" Then
            ws.Range("A1").Value = "Products in the " & cat & " category"
            ws.Range("A1").Font.Bold = True
            ws.Range("A3").Value = "Product"
            ws.Range("B3").Value = "Price"
            ws.Range("A3:B3").Font.Bold = True
            ws.Columns("A:C").AutoFit
            Columns("A").ColumnWidth = 58
        End If
        
        ' Find next empty row, then copy the product as well as the price
        rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(rowCount, 1).Value = cell.Offset(0, -1).Value
        ws.Cells(rowCount, 2).Value = cell.Offset(0, 1).Value
        ws.Cells(rowCount, 2).NumberFormat = "$#,##0.00"
    Next cell
    
    MsgBox "All category worksheets created successfully!", vbInformation
End Sub

' This function makes sure that the new sheet name will be viable, and not caue
' any errors
Function CleanSheetName(sheetName As String) As String
    sheetName = Replace(sheetName, "*", "")
    sheetName = Replace(sheetName, "[", "")
    sheetName = Replace(sheetName, "]", "")
    sheetName = Replace(sheetName, ":", "-")
    sheetName = Replace(sheetName, "/", "-")
    sheetName = Replace(sheetName, "\", "-")
    sheetName = Replace(sheetName, "?", "")
  

    ' Trim to 31 characters as that is the longest length a worksheet can be
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If

    CleanSheetName = sheetName
End Function

' Subroutine to get the width of column A in data sheet so it matches with the data
' in the new category sheet
Sub GetColumnWidth()
    MsgBox "The width of column A is: " & Columns("A").ColumnWidth
End Sub

