VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataIntegratorUserForm 
   Caption         =   "By Alexander Czernik"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11820
   OleObjectBlob   =   "DataIntegratorUserForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "DataIntegratorUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DimTableArray(0 To 100, 0 To 2) As String
Dim originWbPath As String

Private Sub UserForm_Initialize()
    
    Dim wksNames() As String
    
    originWbPath = ActiveWorkbook.path & "\" & ActiveWorkbook.name
    ModeListBox.AddItem cMODCOMPARE
    ModeListBox.AddItem cMODINTEGRATE
    'ModeListBox.AddItem cMODHIGHLIGHT
    
    BaseTableLabel.Caption = ""
    ModeLabel.Caption = ""
    ByKeyOptionsButton.Caption = cOPTBYKEY
    ByKeyOptionsButton.Value = True
    ByAttributeOptionsButton.Caption = cOPTBYATTRIBUTE
    
    BaseWbPathEditText.Text = originWbPath
    BaseWksListBox.Selected(inArray(BaseWksListBox.List, ActiveSheet.name)) = True
    
End Sub

Private Sub ModeListBox_Click()

    If ModeListBox.Selected(0) = True Then
        ModeLabel.Caption = "With"
    ElseIf ModeListBox.Selected(1) = True Then
        ModeLabel.Caption = "From"
    ElseIf ModeListBox.Selected(2) = True Then
    
    Else
        ModeLabel.Caption = ""
    End If
    
    checkIfReady
End Sub

Private Sub BrowseBaseTableButton_Click()
    Dim filePath As String
    
    filePath = browseFilePath()
    BaseWbPathEditText.Text = filePath
End Sub

Private Sub BaseWbPathEditText_Change()
    Dim wb As Workbook
    
    Set wb = getWbReference(BaseWbPathEditText.Text)
    If Not wb Is Nothing Then
        BaseWksListBox.List = getWksNamesInWb(wb)
    Else
        removeItemsfromListBox BaseWksListBox
        deselectItemsfromListBox BaseWksListBox
    End If
    
    removeItemsfromListBox BaseTableListBox
    deselectItemsfromListBox BaseTableListBox
        
    BaseTableLabel.Caption = ""
    checkIfReady
End Sub

Private Sub BaseWksListBox_Click()
    
    If Not IsNull(BaseWksListBox) Then
        wksListBoxWasClicked BaseWbPathEditText, BaseWksListBox, BaseTableListBox
    
        BaseTableLabel.Caption = ""
        checkIfReady
    End If
End Sub

Private Sub BaseTableListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim baseTableName As String
    Dim wksItem As String
    
    If Not IsNull(BaseTableListBox) Then
        baseTableName = BaseTableListBox.List(getSelectedIndex(BaseTableListBox))
        wksItem = BaseWksListBox.List((getSelectedIndex(BaseWksListBox)))
        
        If inArray(OtherTablesListbox.List, baseTableName) = -1 Then
            If Not checkForWks(wksItem, False) Then
                BaseTableLabel.Caption = baseTableName
                checkIfReady
            Else
                MsgBox "Two Tables can't be located on the this Worksheet '" & wksItem & "' .. Sorry "
            End If
        Else
            MsgBox "All Tables must have different names. Please rename Table " & baseTableName
        End If
        deselectItemsfromListBox BaseTableListBox
    End If
End Sub

''''''''''Other Tables'''''''''''''''''''''

Private Sub BrowseOtherTableButton_Click()
    Dim filePath As String
    filePath = browseFilePath()
    OtherWbPathEditText.Text = filePath
End Sub

Private Sub OtherWbPathEditText_Change()
    Dim wb As Workbook
    
    Set wb = getWbReference(OtherWbPathEditText.Text)
    If Not wb Is Nothing Then
        AddOtherWksListBox.List = getWksNamesInWb(wb)
    Else
        removeItemsfromListBox AddOtherWksListBox
        deselectItemsfromListBox AddOtherWksListBox
    End If

    removeItemsfromListBox AddOtherTablesListBox
    deselectItemsfromListBox AddOtherTablesListBox
End Sub

Private Sub AddOtherWksListBox_Click()
    If Not IsNull(AddOtherWksListBox) Then
        wksListBoxWasClicked OtherWbPathEditText, AddOtherWksListBox, AddOtherTablesListBox
    End If
End Sub

Private Sub AddOtherTablesListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim wksItem As String
    Dim newItem As String
    
    If Not IsNull(AddOtherTablesListBox) Then
        newItem = AddOtherTablesListBox.List(getSelectedIndex(AddOtherTablesListBox))
        wksItem = AddOtherWksListBox.List(getSelectedIndex(AddOtherWksListBox))
        If inArray(OtherTablesListbox.List, newItem) = -1 And newItem <> BaseTableLabel.Caption Then
            If Not checkForWks(wksItem, True) Then
                OtherTablesListbox.AddItem newItem
                
                Call appendToTableArray(OtherWbPathEditText.Text, wksItem, _
                   newItem)
            Else
                MsgBox "Two Tables can't be located on the this Worksheet '" & wksItem & "' .. Sorry "
            End If
        Else
            MsgBox "All Tables must have different names. Please rename Table " & newItem
            
        End If
        deselectItemsfromListBox AddOtherTablesListBox
        checkIfReady
    End If
End Sub

Private Sub OtherTablesListbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Integer
    Dim j As Integer
    Dim selectedIndex As Integer
    Dim itemName As String
    
    If Not IsNull(OtherTablesListbox) Then
        selectedIndex = getSelectedIndex(OtherTablesListbox)
        itemName = OtherTablesListbox.List(selectedIndex)
        Call removeFromTableArray(itemName)
        OtherTablesListbox.RemoveItem selectedIndex
        
        deselectItemsfromListBox AddOtherTablesListBox
        
        checkIfReady
    End If
End Sub

Private Sub StartButton_Click()
    Dim tableArray() As String
    Dim tables() As String
    Dim wbs() As String
    Dim wkss() As String

    Dim mode As String
    Dim options(0 To 1) As String        'redim if more options are available
    Dim length As Integer
    
    Dim i As Integer
    
    mode = ModeListBox.List(getSelectedIndex(ModeListBox))
    length = OtherTablesListbox.ListCount
    ReDim tables(0 To length)
    ReDim wbs(0 To length)
    ReDim wkss(0 To length)
    
    tables(0) = BaseTableLabel.Caption
    wbs(0) = BaseWbPathEditText.Text
    wkss(0) = BaseWksListBox.List(getSelectedIndex(BaseWksListBox))
    
    tableArray = returnTableArray()
    For i = 0 To length - 1
        wbs(i + 1) = tableArray(i, 0)
        wkss(i + 1) = tableArray(i, 1)
        tables(i + 1) = tableArray(i, 2)
    Next i
    
    options(0) = mode
    If ByKeyOptionsButton.Value Then
        options(1) = cOPTBYKEY
    Else
        options(1) = cOPTBYATTRIBUTE
    End If
    initDIMaskAndHDITables originWbPath, wbs, wkss, tables, options
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


''''''''''''''''''Useful functions'''''''''''''''''''''''''
Private Function getSelectedIndex(ByRef listBx As Control) As Integer
    Dim i As Integer
    For i = 0 To listBx.ListCount - 1
        If listBx.Selected(i) Then
            getSelectedIndex = i
            Exit Function
        End If
    Next i
    getSelectedIndex = -1
End Function

Private Sub wksListBoxWasClicked(wbEditText As Control, wksListBx As Control, tableListBx As Control)
    Dim wb As Workbook
    Dim i As Integer
    Dim j As Integer
    Dim selectedIndex As Integer
    Dim names() As String
    
    Set wb = getWbReference(wbEditText.Text)
    selectedIndex = getSelectedIndex(wksListBx)
    If Not wb Is Nothing And selectedIndex <> -1 Then
        names = getTableNamesInWks(wb.Worksheets(wksListBx.List(selectedIndex)))
        If Not isNA(names) Then
            tableListBx.List = names
            deselectItemsfromListBox tableListBx
        Else
            removeItemsfromListBox tableListBx
            deselectItemsfromListBox tableListBx
        End If
    Else
        removeItemsfromListBox tableListBx
        deselectItemsfromListBox tableListBx
    End If
End Sub

Private Sub removeItemsfromListBox(listBx As Control)
    Dim i As Integer
    For i = 0 To listBx.ListCount - 1
        listBx.RemoveItem 0
    Next i
End Sub

Private Sub deselectItemsfromListBox(listBx As Control)
    Dim i As Integer
    For i = 0 To listBx.ListCount - 1
        listBx.Selected(i) = False
    Next i
End Sub

Private Sub checkIfReady()
    Dim modeReady As Boolean: modeReady = False
    Dim baseTableReady As Boolean: baseTableReady = False
    Dim otherTableReady As Boolean: otherTableReady = False
    
    Dim i As Integer
    
    For i = 0 To ModeListBox.ListCount - 1
        If ModeListBox.Selected(i) Then
            modeReady = True
        End If
    Next i
    
    If BaseTableLabel.Caption <> "" Then
        baseTableReady = True
    End If
    
    If OtherTablesListbox.ListCount > 0 Then
        otherTableReady = True
    End If
    
    
    If modeReady And baseTableReady And otherTableReady Then
        StartButton.Enabled = True
    Else
        StartButton.Enabled = False
    End If
End Sub

Private Function checkForWks(wksName As String, otherTable As Boolean) As Boolean
    Dim i As Integer
    For i = 0 To 100
        If DimTableArray(i, 1) = wksName Then
            checkForWks = True
            Exit Function
        End If
    Next i
    
    If wksName = BaseWksListBox.List(getSelectedIndex(BaseWksListBox)) And otherTable Then
        checkForWks = True
        Exit Function
    End If
    checkForWks = False
End Function

Private Sub appendToTableArray(wbName As String, wksName As String, tableName As String)
    Dim i As Integer
    For i = 0 To 100
        If IsEmpty(DimTableArray(i, 0)) Or DimTableArray(i, 0) = "" Then
            DimTableArray(i, 0) = wbName
            DimTableArray(i, 1) = wksName
            DimTableArray(i, 2) = tableName
            i = 100
        End If
    Next i
    
End Sub

Private Sub removeFromTableArray(tableName As String)
    Dim i As Integer
    
    For i = 0 To 100
        If DimTableArray(i, 2) = tableName Then
            DimTableArray(i, 0) = ""
            DimTableArray(i, 1) = ""
            DimTableArray(i, 2) = ""
        End If
    Next i
End Sub

Private Function returnTableArray() As String()
    Dim returnArray() As String
    
    Dim counter As Integer
    Dim i As Integer
    
    For i = 0 To 100
        If Not IsEmpty(DimTableArray(i, 0)) And DimTableArray(i, 0) <> "" Then
            counter = counter + 1
        End If
    Next i
    
    ReDim returnArray(0 To counter - 1, 0 To 2)
    counter = 0
    For i = 0 To 100
        If Not IsEmpty(DimTableArray(i, 0)) And DimTableArray(i, 0) <> "" Then
            returnArray(counter, 0) = DimTableArray(i, 0)
            returnArray(counter, 1) = DimTableArray(i, 1)
            returnArray(counter, 2) = DimTableArray(i, 2)
            counter = counter + 1
        End If
    Next i
    
    returnTableArray = returnArray
End Function

