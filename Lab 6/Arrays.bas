Attribute VB_Name = "Lab6"
Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Sushant Sah
' Student ID: 169060628
' Date: 01/03/2025
' Program title: Lab 6 - Modular Programming
' Description: Task 1/2
'===========================================================+

' Function for task 1
Function numExtract(text As String) As Long
    Dim nString As String, character As String
    Dim i As Integer
    
    ' Empty String to append integers to as they come
    nString = ""
    
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        
        ' Append the number to the temp string
        If IsNumeric(character) Then
        nString = nString & character
        End If
    Next i
    
    ' Convert string to long to return
    numExtract = CLng(nString)
    
End Function

'Function for task 2
Function ArrayEqual(array1() As Single, array2() As Single) As Boolean
    
    ' to make loops easier to read and shorter
    Dim i As Integer, UbA1 As Integer, UbA2 As Integer, LbA1 As Integer, LbA2 As Integer
    
    UbA1 = UBound(array1)
    UbA2 = UBound(array2)
    LbA1 = LBound(array1)
    LbA2 = LBound(array2)
    
    
    ' Base case: if the array sizes aren't the same, exit
    If UbA1 - LbA1 <> UbA2 - LbA2 Then
        ArrayEqual = False

        Exit Function
    End If
    
    For i = LbA1 To UbA1
        'If element at the increment point isn't the same, terminate
        If array1(i) <> array2(i) Then
        ArrayEqual = False
        MsgBox "The arrays are not equal!", vbCritical
    
        
        
        Exit Function
        End If
        
    Next i
    
    ' If all elements are the same, then continue
    ArrayEqual = True
    MsgBox "The arrays are equal!", vbInformation
    
    
End Function

' Sub to test task 2
Sub TestMyFunctionsF()
    Dim array1(3) As Single, array2(3) As Single, array3(3) As Single
    Dim array4(3) As Single
    
    Dim res As Boolean, res2 As Boolean
    
    ' Hardcoding the array values to test
    array1(0) = 2
    array1(1) = 5
    array1(2) = 2
    array1(3) = 9
    
    array2(0) = 4
    array2(1) = 3
    array2(2) = 6
    array2(3) = 2
    
    ' Calling "ArrayEqual" to test the two arrays for equality
    res = ArrayEqual(array1, array2)
    
    ' Outputting results to the excel sheet
    Range("B10").Value = IIf(res, "TRUE", "FALSE")
        
End Sub

' Sub to test task 2
Sub TestMyFunctionsT()
    Dim array1(3) As Single, array2(3) As Single
    
    Dim res As Boolean
    
    ' Hardcoding the array values to test
    array1(0) = 5
    array1(1) = 2
    array1(2) = 5
    array1(3) = 2
    
    array2(0) = 5
    array2(1) = 2
    array2(2) = 5
    array2(3) = 2
    
    
    ' Calling "ArrayEqual" to test the two arrays for equality
    res = ArrayEqual(array1, array2)
    
    ' Outputting results to the excel sheet
    Range("B11").Value = IIf(res, "TRUE", "FALSE")
        
End Sub

' Sub to clear contents to re-test
Sub clearText()
    Range("B10:B11").ClearContents

End Sub

