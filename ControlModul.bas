Attribute VB_Name = "ControlModul"
Option Explicit

Sub initDIMaskAndHDITables(originWbPath As String, wbNames() As String, wksNames() As String, tableNames() As String, options() As String)  'is being called from userform
    Dim wb As Workbook: Set wb = getWbReference(originWbPath)
    Dim data As Variant
    
    If Not getWksByName(wb, "DI Mask") Is Nothing Then
        MsgBox "You have already created a DI Mask. Please first delete the current DI Mask."
    End If
    
    deleteHDITables wb
    
    data = getData(wbNames, wksNames, tableNames)
    
    filDIMask wb, data, wksNames, tableNames, options 'ATTENTION has to be wb with di mask
    filHDITables wb, data, wbNames, wksNames, tableNames 'ATTENTION has to be wb reference to hiddenwb (wb in data integrator)

    wb.Activate
    
    wb.Worksheets("DI Mask").Select   'ATTENTION has to be wb with di mask
End Sub

Sub StartDI() 'activates when start button on di mask wks was pushed
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wks As Worksheet
    Dim integrationRef() As String
    Dim dataModel As dataModel
    
    Dim mode As String
    Dim options() As String
    
    Set wks = getWksByName(wb, "DI Mask")
    If wks Is Nothing Then
        MsgBox "Please do not change the name of the Worksheet 'DI Mask'"
        onBug
        End
    End If
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = False
    End With
    
    Set dataModel = initDIMaskModel(wb, wks)
    options = Split(wks.Range(cMODE), " ")

    If options(0) = cMODCOMPARE Then
        ModeCompare wb, dataModel, options
    ElseIf options(0) = cMODINTEGRATE Then
        
        ModIntegrate wb, dataModel, options, getTableReference(wb, 0)
    ElseIf options(0) = cMODHIGHLIGHT Then
    
    End If

    onBug
End Sub

Function initDIMaskModel(wb As Workbook, wks As Worksheet) As dataModel

    Dim listObj As ListObject
    
    Dim dataModel As New dataModel
    Dim diMasks() As DIMaskModel
    Dim diMask As DIMaskModel
    Dim counter As Integer
    
    Dim columnIndex As Integer
    Dim keyIndex As Integer
    Dim attributeIndex As Integer

    If validateDIMaskWks(wks) Then
        With wks
            columnIndex = .ListObjects(1).ListColumns(cCOLUMNCOLUMN & " " & 0).index - 1
            keyIndex = .ListObjects(1).ListColumns(cKEYCOLUMN & " " & 0).index - 1
            attributeIndex = .ListObjects(1).ListColumns(cATTRIBUTECOLUMN & " " & 0).index - 1
            ReDim diMasks(0 To .ListObjects.Count - 1)
            For Each listObj In .ListObjects
                With listObj
                    Set diMask = New DIMaskModel
                    diMask.init .name, columnIndex, keyIndex, attributeIndex, arraysStartWith0(.DataBodyRange.Value2)
                    Set diMasks(counter) = diMask
                    counter = counter + 1
                End With
            Next listObj
        End With
    Else
        onBug
        End
    End If
    
    dataModel.init diMasks
    
    Set initDIMaskModel = dataModel
End Function

