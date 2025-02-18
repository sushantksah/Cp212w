Attribute VB_Name = "Task1"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: January 31st, 2025
' Program title: Task 1
' Description: Title
'===========================================================+

Sub FormatTitle()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Naming Ranges
    ws.Range("A1").Name = "Title"
    ws.Range("B2:H2").Name = "Headings"
    ws.Range("A3:A33").Name = "TeamsSequence"
    
    ' Formatting Ranges & Adding Color
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 16
        End With
    
    With ws.Range("A1:H1")
        .Merge
        .HorizontalAlignment = xlCenter
        End With
        
    With ws.Range("Headings")
        .Font.Bold = True
        .Font.Italic = True
        .Font.Color = RGB(0, 0, 255)
        End With
        
    With ws.Range("TeamsSequence")
        .Font.Italic = True
        .Font.Color = RGB(165, 50, 51)
        End With
        
    ' Average in B34
    ws.Range("C34").Formula = "=ROUND(Average(C3:C33),2)"
    ws.Range("B34").Value = "Average"
    ws.Range("B34").Font.Bold = True
    
    ' Max in B35
    ws.Range("B35").Value = "Max"
    ws.Range("B35").Font.Bold = True
    ws.Range("D35").Formula = "=MAX(D3:D33)"
    ws.Range("D35").Copy Destination:=ws.Range("E35:F35")

    
    ' Formula for Total Points and Points Percentage
    ws.Range("G3").Formula = "=(D3*2) + F3"
    ws.Range("H3").Formula = "=ROUND(G3 / (C3*2),2)"
    
    ' Setting Formulas for Ranges
    ws.Range("G3").Copy Destination:=ws.Range("G4:G33")
    ws.Range("H3").Copy Destination:=ws.Range("H4:H33")

    ' Extra Stuff
    ' Wasn't sure if it meant to just Average the averages (all GP is 56), so I extended it to the other data points
    ws.Range("C34").Copy Destination:=ws.Range("D34:H34")
    
End Sub

' Clearing All Formatting
Sub ClearRoutine()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Delete the Range Names Set
    ThisWorkbook.Names("Title").Delete
    ThisWorkbook.Names("Headings").Delete
    ThisWorkbook.Names("TeamsSequence").Delete
    
    ' Reset Formatting
    ws.Range("G3:H33").ClearFormats
    ws.Range("A34:H35").ClearFormats
    ws.Range("G3:H33").ClearContents
    ws.Range("A34:H35").ClearContents
    ws.Range("B2:H2").ClearFormats
    ws.Range("A1:H1").ClearFormats
    ws.Range("A3:A33").ClearFormats
    
End Sub
    
    
    
        

