Attribute VB_Name = "RF1"
Option Explicit

Sub DeleteWs()
    Dim ws As Worksheet
    
    ' Find/Delete "RF1" by looping through all worksheets
    For Each ws In Worksheets
        If ws.Name = "RF1" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            
            MsgBox "Worksheet deleted successfully!", vbInformation
            Exit Sub
           End If
          Next ws
            
    ' If "RF1" isn't found
    MsgBox "Worksheet not found!", vbInformation
End Sub

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 07/02/2025
' Program title: RF1
' Description: Making new worksheet
'===========================================================+

Sub MakeRF1()
    Dim ws As Worksheet, newWs As Worksheet
    Dim rng As Range, cell As Range
    Dim lastRow As Integer
    
    ' Check/Exit if already on RF1
    For Each ws In Worksheets
        If ws.Name = "RF1" Then
            Exit Sub
        End If
        
    Next ws
    
    ' Make new worksheet called "RF1"
    Set newWs = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    newWs.Name = "RF1"
    
    ' Headers for "RF1"
    newWs.Cells(1, 1).Value = "Worksheet"
    newWs.Cells(1, 2).Value = "Formula"
    newWs.Cells(1, 3).Value = "Value"
    newWs.Rows(1).Font.Bold = True
    lastRow = 2
    
    ' Loop through worksheets to look for sells with formulas
    For Each ws In Worksheets
        If ws.Name <> "RF1" Then
            Set rng = ws.UsedRange
   
            For Each cell In rng
                If cell.HasFormula Then
                    newWs.Cells(lastRow, 1).Value = ws.Name
                    newWs.Cells(lastRow, 2).Value = "'" & cell.Formula
                    newWs.Cells(lastRow, 3).Value = cell.Value
                    lastRow = lastRow + 1
                End If
                
            Next cell
    
            
        End If
        
    Next ws
    
    ' Widening the colummns to make sure the words have enough room to fit
    newWs.Columns("A:C").AutoFit
    
    'If everything works properly, message will display
    MsgBox "Worksheet created successfully!", vbInformation
End Sub

    
    

