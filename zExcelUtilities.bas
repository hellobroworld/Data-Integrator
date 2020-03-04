Attribute VB_Name = "zExcelUtilities"
Option Explicit

Sub formatCells(wks As Worksheet, row As Integer, col As Integer, mode As Integer)

    With wks.Cells(row, col)
        If mode = 1 Then        'Überschrift 1
            .Font.Size = 16
        ElseIf mode = 2 Then
            .Font.Size = 14
        ElseIf mode = 3 Then
            .Font.Size = 12
        End If
        
        If mode < 4 Then
            .Font.Bold = True
        End If
    End With
End Sub
