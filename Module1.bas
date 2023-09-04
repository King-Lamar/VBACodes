Attribute VB_Name = "Module1"
' Copyright (c) Michael Jimoh, 2023
' dev.jmichael@gmail.com
' All rights reserved.


Sub SeparateNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim fullName As String
    Dim nameParts() As String
    Dim middleName As String
    Dim lastName As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("prognosis_Master") ' Change the sheet name as needed

    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' Loop through the data starting from row 2
    For i = 2 To lastRow
        ' Use CStr to convert cell value to a string
        fullName = Trim(CStr(ws.Cells(i, "F").Value)) ' Extract data from column F

        ' Split the full name into parts
        nameParts = Split(fullName, " ")

        ' Initialize last name and middle name
        lastName = ""
        middleName = ""

        ' Determine the last name based on the number of parts
        If UBound(nameParts) >= 1 Then
            lastName = nameParts(UBound(nameParts))
        End If

        ' Determine the middle name if there are three parts
        If UBound(nameParts) >= 2 Then
            For j = 1 To UBound(nameParts) - 1
                middleName = middleName & nameParts(j) & " "
            Next j
            middleName = Trim(middleName)
        End If

        ' Place the last name in the "Last Name" column and middle name in the "Middle Name" column
        ws.Cells(i, "H").Value = lastName
        ws.Cells(i, "I").Value = middleName
    Next i
End Sub

