Attribute VB_Name = "FilModul"
Option Explicit

Public Const cKEYCOLUMN As String = "Key"
Public Const cATTRIBUTECOLUMN As String = "Attribute"
Public Const cCOLUMNCOLUMN As String = "Spalten"

Public Const cMODE As String = "Mode"

Const cSTARTROW As Integer = 1
Const cSTARTCOL As Integer = 1
Const cSPACES As Integer = 2

Sub filDIMask(wb As Workbook, data As Variant, wksPaths() As String, tableNames() As String, options() As String)
    Dim wks As Worksheet
    Dim listObj As ListObject
    Dim shapeObj As Shape
    
    Dim rangeArray As Variant
    Dim header As Variant
    
    Dim length As Integer
    Dim lengthHeader As Integer
    Dim i As Integer
    
    Dim optionsString As String
    
    Dim row As Integer
    Dim col As Integer
    length = UBound(tableNames)
    
    Set wks = wb.Sheets.Add
    With wks
        .Cells.Font.name = "Arial"
        .Cells.Font.Size = 12
        .name = "DI Mask"
        .Cells(cSTARTROW, cSTARTCOL) = "Data Integrator"
        formatCells wks, cSTARTROW, cSTARTCOL, 1

        'input mode and options
        optionsString = options(0)
        For i = 1 To UBound(options)
            optionsString = optionsString & " " & options(i)
        Next i
        With .Cells(cSTARTROW + cSPACES, cSTARTCOL + cSPACES)
            .Value2 = optionsString
            .name = cMODE
        End With
        formatCells wks, cSTARTROW + cSPACES, cSTARTCOL + cSPACES, 2
        '---
        
        Set shapeObj = createStartShape(wks)
        shapeObj.OnAction = "StartDI"
        
        For i = 0 To length
            rangeArray = data(i)
            header = convertArrayToRange(getHeaderFromMatrix(rangeArray))       'problem gibt 2 arrays zurück
            lengthHeader = UBound(header, 1) - 1
            row = cSTARTROW + 2 * cSPACES
            col = cSTARTCOL + cSPACES + i * 4
            
            .Cells(row, col) = "Tabelle: " & wksPaths(i) & " Tabelle: " & tableNames(i)
            formatCells wks, row, col, 3
            
            .Cells(row + 1, col) = cCOLUMNCOLUMN & " " & i
            .Cells(row + 1, col).ColumnWidth = 30
            With .Range(.Cells(row + 2, col), .Cells(row + 2 + lengthHeader, col))
                .Value = header
                .Interior.Color = RGB(111, 111, 111)
                .Font.Color = RGB(255, 255, 255)
                .WrapText = True
            End With
            .Cells(row + 1, col + 1) = cKEYCOLUMN & " " & i
            .Cells(row + 1, col + 2) = cATTRIBUTECOLUMN & " " & i
            Set listObj = .ListObjects.Add(xlSrcRange, .Range(.Cells(row + 1, col), .Cells(row + 2 + lengthHeader, col + 2)), False, xlYes)
            listObj.name = "DI " & tableNames(i)        'DI: data integrator
        Next i
    End With
End Sub

Sub filHDITables(wb As Workbook, data As Variant, wbPaths() As String, wksPaths() As String, tableNames() As String)
    Dim wks As Worksheet
    Dim listObj As ListObject
    Dim i As Integer
    Dim length As Integer
    Dim width As Integer
    
    For i = 0 To UBound(data)
        Set wks = wb.Worksheets.Add
        length = UBound(data(i), 1) + 1
        width = UBound(data(i), 2) + 1
        With wks
            .name = "HDI " & wksPaths(i)   'HDI: hidden data integrator
            .Range(.Cells(1, 1), .Cells(length, width)).Value2 = data(i)
            
            Set listObj = .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(length, width)), False, xlYes)
            listObj.name = "HDI " & tableNames(i)
            
            .Visible = xlSheetVeryHidden
        End With
    Next i
    
    Set wks = wb.Worksheets.Add
    length = UBound(wbPaths)
    width = UBound(wksPaths)
    
    With wks
        nameWks wb, wks, cHIDDENWKS
        .Cells(1, 1).Value2 = "Wb_Path"
        .Cells(1, 2).Value2 = "Wks_Name"
        .Cells(1, 3).Value2 = "Table_Name"
        
        .Range(.Cells(2, 1), .Cells(2 + length, 1)) = wbPaths
        .Range(.Cells(2, 2), .Cells(2 + length, 2)) = wksPaths
        .Range(.Cells(2, 3), .Cells(2 + length, 3)) = tableNames
        
        Set listObj = .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(2 + length, 3)), False, xlYes)
        listObj.name = cHIDDENTABLE
        .Visible = xlSheetVeryHidden
    End With
    
End Sub

Function createStartShape(wks As Worksheet) As Shape
    Dim shapeObj As Shape
    
    Set shapeObj = wks.Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 12, 138.75, 40.5)
    
    With shapeObj
        .Fill.UserPicture "H:\data\Downloads\Matrixbild.jpg"

        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.name = "Times New Roman"
                .Font.Size = 20
                .Font.Bold = True
                .Font.Fill.ForeColor.RGB = 16777215
                .Characters.Text = "Start"
            End With
        End With
    End With
    
    Set createStartShape = shapeObj
End Function



