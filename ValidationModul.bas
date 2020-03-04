Attribute VB_Name = "ValidationModul"
Option Explicit

Function validateDIMaskWks(wks As Worksheet) As Boolean 'columnIndices = ("ColName", index)
    Dim listObj As ListObject
    
    Dim validationArray As Variant
    Dim validationHeader As Variant
    Dim validationAttributColumn As Variant
    
    Dim compareArray As Variant
    Dim compareHeader As Variant
    
    Dim i As Integer, j As Integer
    Dim width As Integer
    
    With wks
        For i = 1 To .ListObjects.Count
            validationArray = arraysStartWith0(.ListObjects(i).DataBodyRange.Value2)
            validationHeader = arraysStartWith0(.ListObjects(i).HeaderRowRange.Value2)
            
            If Not onlyUniqueValuesInTable(validationArray, validationHeader) Then
                validateDIMaskWks = False
                Exit Function
            End If
            validationAttributColumn = getColumnFromMatrix(validationArray, inArray(validationHeader, cATTRIBUTECOLUMN & " " & i - 1))
            For j = 1 To .ListObjects.Count
                If i <> j Then
                    compareArray = arraysStartWith0(.ListObjects(j).ListColumns(cATTRIBUTECOLUMN & " " & j - 1).DataBodyRange.Value2)
                    If Not doItemsInArrayMatch(compareArray, validationAttributColumn) Then
                        MsgBox "Please use for each attribute in " & .ListObjects(j).name & " the same name as in " & .ListObjects(i).name
                        validateDIMaskWks = False
                        Exit Function
                    End If
                End If
            Next j
        Next i
    End With
    
    validateDIMaskWks = True
End Function

Private Function onlyUniqueValuesInTable(validationArray As Variant, validationHeader As Variant) As Boolean
    Dim columnOfArray As Variant
    Dim k As Integer
    For k = 0 To UBound(validationArray, 2)
        columnOfArray = removeNARowsFromArray(getColumnFromMatrix(validationArray, k))
        If Not IsEmpty(columnOfArray) Then
            If UBound(columnOfArray) > 0 And Split(validationHeader(k), " ")(0) = cKEYCOLUMN Then
                MsgBox "Please use only one key in " & validationHeader(k)
                onlyUniqueValuesInTable = False
                Exit Function
            ElseIf UBound(columnOfArray) <> UBound(uniqueInArray(columnOfArray)) Then
                MsgBox "Please ensure that you use different names for all cell values in " & validationHeader(k) & Chr(10) & _
                    "In case the " & cCOLUMNCOLUMN & " values have the same name. Rename these columns and restart the macro."
                onlyUniqueValuesInTable = False
                Exit Function
            End If
        ElseIf Not Split(validationHeader(k), " ")(0) = cATTRIBUTECOLUMN Then
            MsgBox "Please ensure that your column: " & validationHeader(k) & " has a value"
            onlyUniqueValuesInTable = False
            Exit Function
        End If
    Next k
    
    onlyUniqueValuesInTable = True
End Function


Sub linkMe(wks As Worksheet, row As Integer, col As Integer)
    
    wks.Cells(row, col).Value2 = "Visit: 'https://github.com/hellobroworld' for updates and other useful tools."
    
End Sub

