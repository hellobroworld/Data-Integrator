Attribute VB_Name = "DataModul"
Option Explicit

Public Const cHIDDENWKS As String = "HDI hidden paths"
Public Const cHIDDENTABLE As String = "HDI hidden table"

Sub deleteHDITables(wb As Workbook)
    Dim wks As Worksheet
    Dim nameArray() As String
    
    For Each wks In wb.Worksheets
        nameArray = Split(wks.name, " ")
        If nameArray(0) = "HDI" Then
            Application.DisplayAlerts = False
            wks.Visible = xlSheetHidden
            wks.Delete
            Application.DisplayAlerts = False
        End If
    Next wks
End Sub

Function getData(wbPaths() As String, wksNames() As String, tableNames() As String) As Variant
    Dim wb As Workbook
    Dim wks As Worksheet
    Dim tableObj As ListObject
    
    Dim length As Integer
    Dim i As Integer
    Dim rangeArray As Variant
    Dim returnArray As Variant
    
    length = UBound(wbPaths)
    
    ReDim returnArray(0 To length)
    For i = 0 To length
        Set wb = getWbReference(wbPaths(i))
        Set wks = wb.Sheets(wksNames(i))
        Set tableObj = wks.ListObjects(tableNames(i))
        
        returnArray(i) = arraysStartWith0(tableObj.Range.Value2)
        'wb.Close False     'Problem closes all wb.. uncomment to boost performance
    Next i
    
    getData = returnArray
End Function

Function retrieveHDIData(wb As Workbook, keyPairObj As keyPair) As Variant
    Dim returnArray As Variant
    Dim wksObj As Worksheet
    Dim listObj As ListObject

    Dim length As Integer: length = UBound(keyPairObj.compareTableNames, 1)
    Dim compareTableNames() As String
    
    Dim i As Integer
    
    ReDim returnArray(0 To length)
    
    compareTableNames = keyPairObj.compareTableNames
    
    For i = 0 To length
        For Each wksObj In wb.Worksheets            'Search for hdi-table
            For Each listObj In wksObj.ListObjects
                If i = 0 Then
                    If listObj.name = "H" & keyPairObj.baseTableName Then
                    
                        returnArray(0) = arraysStartWith0(listObj.Range.Value2)
                    End If
                Else
                    If listObj.name = "H" & compareTableNames(i) Then
                    
                        returnArray(i) = arraysStartWith0(listObj.Range.Value2)
                    End If
                End If
            Next listObj
        Next wksObj
    Next i
    
    retrieveHDIData = returnArray
End Function

Function getTableReference(wb As Workbook, tableIndex As Integer) As String()    '{wb path, wks name, table name}
    Dim returnArray(0 To 2) As String
    Dim table As Variant
    Dim i As Integer
    
    table = arraysStartWith0(wb.Sheets(cHIDDENWKS).ListObjects(cHIDDENTABLE).DataBodyRange.Value2)
    
    For i = 0 To 2
        returnArray(i) = table(tableIndex, i)
    Next i
    
    getTableReference = returnArray
End Function
