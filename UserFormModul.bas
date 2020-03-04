Attribute VB_Name = "UserFormModul"
Option Explicit

Public Const cMODCOMPARE As String = "Compare"
Public Const cMODINTEGRATE As String = "Integrate"
Public Const cMODHIGHLIGHT As String = "Highlight"

Public Const cOPTBYKEY As String = "Key-Dont-Match"
Public Const cOPTBYATTRIBUTE As String = "Key-Match"

Sub ActivateDIUserform()
    Dim wks As Worksheet
    Dim response As String
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = False
    End With
    
    For Each wks In ActiveWorkbook.Sheets
        If wks.name = "DI Mask" Then
            response = MsgBox("You first have to delete the current DI Mask." & Chr(10) & "Do you want to delete it?", vbYesNo, "Before you start")
            If response = vbNo Then
                onBug
                End
            End If
            wks.Delete
            deleteHDITables ActiveWorkbook
        End If
    Next wks
    
    response = MsgBox("Have you ensured that non of your objects(Worksheets, Tables) do start with 'DI' or 'HDI'?", vbYesNo, "Before you start")
    
    If response = vbYes Then
        DataIntegratorUserForm.Show
    End If
    
    onBug
End Sub


Public Sub onBug()
With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With
End Sub
