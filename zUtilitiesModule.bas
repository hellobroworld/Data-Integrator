Attribute VB_Name = "zUtilitiesModule"


'FILE OPERATIONS
Function FileExists(filePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Function browseFilePath() As String
On Error GoTo err
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Any file is selected
            browseFilePath = .SelectedItems.Item(1)
        Else ' else dialog is cancelled
            MsgBox "You have cancelled the dialogue"
        End If
    End With
err:
    Exit Function
End Function

Function IsWorkBookOpen(fileName As String)
    Dim file As Long
    Dim ErrNo As Long

    On Error Resume Next
    file = FreeFile()
    Open fileName For Input Lock Read As #file
    Close file
    ErrNo = err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Function WorksheetExists(wksName As String) As Boolean

For Each wks In Worksheets
    If wks.name = wksName Then
        WorksheetExists = True
        Exit Function
    End If
Next wks

WorksheetExists = False
End Function
'ONLY AVAILABLE IF POWERPOINT ADD IN IS CHECKED
'Function openPPFile(pp As PowerPoint.Application, filePath As String) As PowerPoint.Presentation
'On Error GoTo err
    
  '  Set openPPFile = pp.Presentations.Open(filePath)
  '  Exit Function
'err:
'MsgBox "Please select valide File"
'End
'End Function

Function getWbReference(path As String) As Workbook
    Dim wb As Workbook
    Dim wbObj As Workbook
    Dim wbName As String
    
    wbName = getFileName(path)
    
    For Each wbObj In Workbooks
        If wbObj.name = wbName Then
            Set wb = wbObj
        End If
    Next wbObj
    
    If wb Is Nothing Then
        On Error GoTo noWbFound
        Set wb = Workbooks.Open(path)
        On Error GoTo 0
    End If
    
    Set getWbReference = wb
    
noWbFound:
Exit Function
End Function

Function getFileName(path As String) As String
    Dim nameArray() As String
    nameArray = Split(path, "\")
    getFileName = nameArray(UBound(nameArray))
End Function

Function getWksByName(wb As Workbook, name As String) As Worksheet
    Dim wks As Worksheet
    
    For Each wks In wb.Worksheets
        If wks.name = name Then
            Set getWksByName = wks
            Exit Function
        End If
    Next wks
End Function

Sub nameWks(wb As Workbook, wks As Worksheet, name As String)
    
    Dim otherWks As Worksheet: Set otherWks = getWksByName(wb, name)
    Dim stringArray() As String
    Dim newName As String
    Dim firstPart As Integer
    Dim i As Integer
    Dim index As Integer
    
    If otherWks Is Nothing Then
        wks.name = name
    Else
        stringArray = Split(otherWks.name, " ")
        firstPart = check(stringArray(0), "Integer")
        If firstPart = 0 Then
            newName = "1"
            index = 0
        Else
            firstPart = firstPart + 1
            newName = firstPart
            index = 1
        End If
        
        For i = index To UBound(stringArray)
            newName = newName & " " & stringArray(i)
        Next i
        nameWks wb, wks, newName
    End If

End Sub


Function getWksNamesInWb(wb As Workbook) As String()
    Dim wks As Worksheet
    Dim length As Integer
    Dim returnArray() As String
    Dim counter As Integer
    
    ReDim returnArray(0 To wb.Worksheets.Count - 1)
    For Each wks In wb.Worksheets
        returnArray(counter) = wks.name
        counter = counter + 1
    Next wks
    
    getWksNamesInWb = returnArray
End Function

Function getTableNamesInWks(wks As Worksheet) As String()
    Dim listObj As ListObject
    Dim length As Integer
    Dim returnArray() As String
    Dim counter As Integer
    If Not wks.ListObjects.Count = 0 Then
        ReDim returnArray(0 To wks.ListObjects.Count - 1)
        For Each listObj In wks.ListObjects
            returnArray(counter) = listObj.name
            counter = counter + 1
        Next listObj
    End If
    getTableNamesInWks = returnArray
End Function
'------------------------------------------------------------------------------------------------------------------------------------

Function check(dataType As Variant, desiredDataType As String) As Variant

    Select Case desiredDataType
        Case "Integer"
            On Error Resume Next
            check = CInt(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = 0
            End If
        Case "Single"
            On Error Resume Next
            check = CSng(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = 0
            End If
        Case "Double"
            On Error Resume Next
            check = CDbl(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = 0
            End If
        Case "String"
            On Error Resume Next
            check = CStr(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = ""
            End If
    End Select

End Function
'-----------------------------------------------------------------------------------------------------------------------

''------------------------------------------------------------------------------------------------------------------------------------
'String OPERATIONS
Function removeWhiteSpaces(oString As String) As String
length = Len(oString)

For i = 1 To length
    iChar = Mid(oString, i, 1)
    If iChar <> " " Then
        nString = nString & iChar
    End If
Next i

removeWhiteSpaces = nString
End Function

Function removeLinebreaks(myString As String) As String
    Dim arrayString() As String
    Dim newString As String
    'newString = ""
    For i = 1 To Len(myString) Step 2
        
        iChar = Mid(myString, i, 2)
        If iChar <> vbCrLf And iChar <> vbNewLine And iChar <> Chr(10) And iChar <> Chr(13) Then
            newString = newString & iChar
        End If
    Next i
    
    removeLinebreaks = newString
End Function

Function removeFirstPartFromString(myString As String) As String
    Dim splitString() As String
    Dim newString As String
    splitString = Split(myString, " ")
    
    newString = splitString(1)
    For i = 2 To UBound(splitString)
        newString = " " & newString
    Next i
    
    removeFirstPartFromString = newString
End Function


'------------------------------------------------------------------------------------------------------------------------------------
'ARRAY OPERATIONS
    'standarize Array +++++++++++++++++++++++++++++++++++++++++++++
Function arraysStartWith0(arr As Variant) As Variant
    Dim startLength As Integer
    Dim startWidth As Integer
    Dim length As Integer
    Dim width As Integer
    Dim i As Integer
    Dim j As Integer
    Dim returnArray As Variant
    
    length = UBound(arr, 1)
    width = UBound(arr, 2)
    startLength = LBound(arr, 1)
    startWidth = LBound(arr, 2)
    
    If width - startWidth <> 0 And length - startLength <> 0 Then
        ReDim returnArray(0 To length - startLength, 0 To width - startWidth)
        
        For i = startLength To length
            For j = startWidth To width
                returnArray(i - startLength, j - startWidth) = arr(i, j)
            Next j
        Next i
    ElseIf width - startWidth <> 0 Then
        ReDim returnArray(0 To width - startWidth)
        
        For i = startWidth To width
            For j = startLength To length
                returnArray(i - startWidth) = arr(j, i)
            Next j
        Next i
    Else
        ReDim returnArray(0 To length - startLength)
        
        For i = startLength To length
            For j = startWidth To width
                returnArray(i - startLength) = arr(i, j)
            Next j
        Next i
    End If
    arraysStartWith0 = returnArray
End Function

Function convertArrayToRange(arr As Variant) As Variant
    Dim returnArray As Variant
    
    length = UBound(arr)
    
    ReDim returnArray(1 To length + 1, 1 To 1)
    
    For i = 0 To length
        returnArray(i + 1, 1) = arr(i)
    Next i
    
    convertArrayToRange = returnArray
End Function
    'NA Operations +++++++++++++++++++++++++++++++++++++++++++++
Function isNA(arr As Variant) As Boolean
On Error GoTo NA

    x = UBound(arr)
    If x = -1 Then
        isNA = True
    Else
        isNA = False
    End If
    Exit Function

NA:
On Error GoTo 0
isNA = True
End Function

Function removeNARowsFromMatrix(arr As Variant) As Variant
    Dim length As Integer
    Dim width As Integer
    Dim i As Integer
    Dim j As Integer
    Dim naRowCounter As Integer
    Dim rowIsNA As Boolean: rowIsNA = True
    Dim returnArray As Variant

    length = UBound(arr, 1)
    width = UBound(arr, 2)
    
    For i = 0 To length
        For j = 0 To width
            If Not IsEmpty(arr(i, j)) Then
                rowIsNA = False
            End If
        Next j
        If rowIsNA Then
            naRowCounter = naRowCounter + 1
        End If
        rowIsNA = True
    Next i
    
    If length < naRowCounter Then
        returnArray = Empty
    Else
        ReDim returnArray(0 To length - naRowCounter, 0 To width)
        naRowCounter = 0
        For i = 0 To length
            For j = 0 To width
                If Not IsEmpty(arr(i, j)) Then
                    rowIsNA = False
                End If
            Next j
            If Not rowIsNA Then
                filArray returnArray, naRowCounter, arr, i
                naRowCounter = naRowCounter + 1
            End If
            rowIsNA = True
        Next i
    End If
    removeNARowsFromMatrix = returnArray
End Function

Function removeNARowsFromArray(arr As Variant) As Variant
    Dim length As Integer
    Dim i As Integer
    Dim naRowCounter As Integer
    Dim rowIsNA As Boolean: rowIsNA = True
    Dim returnArray As Variant

    length = UBound(arr, 1)
    
    For i = 0 To length
        
        If Not IsEmpty(arr(i)) Then
            rowIsNA = False
        End If
        If rowIsNA Then
            naRowCounter = naRowCounter + 1
        End If
        rowIsNA = True
    Next i
    
    If length < naRowCounter Then
        returnArray = Empty
    Else
        ReDim returnArray(0 To length - naRowCounter)
        naRowCounter = 0
        For i = 0 To length
            
            If Not IsEmpty(arr(i)) Then
                rowIsNA = False
            End If

            If Not rowIsNA Then
                returnArray(naRowCounter) = arr(i)
                naRowCounter = naRowCounter + 1
            End If
            rowIsNA = True
        Next i
    End If
    removeNARowsFromArray = returnArray
End Function

Function colIsNA(arr As Variant, col As Integer) As Boolean
    Dim i As Integer
    
    For i = 0 To UBound(arr)
        If Not IsEmpty(arr(i, col)) Then
            isColNA = False
            Exit Function
        End If
    Next i
    
    colIsNA = True
End Function
    
    'array transformation +++++++++++++++++++++++++++++++++++++++++++++
Function reDimNxNArray(arr As Variant, sizeL As Integer, sizeW As Integer) As Variant
    Dim returnArray As Variant
    
    Dim width As Integer
    Dim length As Integer
    Dim width1 As Integer
    Dim length1 As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    If IsEmpty(arr) Then
        ReDim returnArray(0 To sizeL, 0 To sizeW)
    Else
        length1 = UBound(arr, 1)
        width1 = UBound(arr, 2)
        If sizW > width1 Then
            width = width1
        Else
            width = sizeW
        End If
        If sizeL > length1 Then
            length = length1
        Else
            length = sizeL
        End If
        ReDim returnArray(0 To sizeL, 0 To sizeW)
        For i = 0 To length
            For j = 0 To width
                returnArray(i, j) = arr(i, j)
            Next j
        Next i
    End If
    
    reDimNxNArray = returnArray
End Function

Function emptyArray(arr As Variant) As Variant
    Dim i As Integer
    
    For i = 0 To UBound(arr)
        arr(i) = Emtpy
    Next i
    
    emptyArray = arr
End Function

    'reference operation +++++++++++++++++++++++++++++++++++++++++++++
Function getIndexOfCol(arr As Variant, colName As String) As Integer
    Dim width As Integer
    width = UBound(arr, 2)
    
    For j = 0 To width
        If arr(0, j) = colName Then
            getIndexOfCol = j
            Exit Function
        End If
    Next j

End Function

Function getHeaderFromMatrix(arr As Variant) As String()
    Dim width As Integer
    Dim j As Integer
    Dim returnArray() As String
    
    width = UBound(arr, 2)
    ReDim returnArray(0 To width)
    For j = 0 To width
        returnArray(j) = arr(0, j)
    Next j
    
    getHeaderFromMatrix = returnArray
End Function

Function getColumnFromMatrix(arr As Variant, col As Integer) As Variant
    Dim length As Integer: length = UBound(arr)
    Dim i As Integer
    
    Dim returnArray As Variant

    ReDim returnArray(0 To length)
    
    For i = 0 To length
        returnArray(i) = arr(i, col)
    Next i
    getColumnFromMatrix = returnArray
End Function

    'input operation +++++++++++++++++++++++++++++++++++++++++++++
Sub filArray(ByRef arr1 As Variant, index1 As Integer, arr2 As Variant, index2 As Integer)
    For j = LBound(arr1, 2) To UBound(arr1, 2)
        arr1(index1, j) = arr2(index2, j)
    Next j
End Sub

Sub addHeaderToArray(ByRef arr As Variant, headerNames As Variant)
    Dim returnArray As Variant
    Dim length As Integer
    Dim width As Integer
    length = UBound(arr, 1)
    width = UBound(arr, 2)
    
    ReDim returnArray(0 To length + 1, 0 To width)
    
    For i = 0 To length + 1
        For j = 0 To width
            If i = 0 Then
                returnArray(0, j) = headerNames(j)
            Else
                returnArray(i, j) = arr(i - 1, j)
            End If
        Next j
    Next i
    
    arr = returnArray
End Sub

    'connection operations +++++++++++++++++++++++++++++++++++++++++++++
'Input: Array of Matrices you want to connect -> Array(mat1, mat2); Output: connected Matrix
Function connectMatrix(arrs As Variant) As Variant
Dim returnArray As Variant
Dim width As Integer
Dim widCounter As Integer
length = UBound(arrs(0), 1)

For Each arr In arrs
    width = width + UBound(arr, 2) + 1
Next arr

ReDim returnArray(0 To length, 0 To width - 1)

For Each arr In arrs
    For i = 0 To length
        width = UBound(arr, 2)
        For j = 0 To width
            returnArray(i, widCounter + j) = arr(i, j)
        Next j
    Next i
    widCounter = widCounter + width + 1
Next arr

connectMatrix = returnArray
End Function

Function connectArrays(arrs As Variant) As Variant
Dim returnArray As Variant
Dim width As Integer
width = leng(arrs)
length = leng(arrs(0))
ReDim returnArray(0 To length, 0 To width)

For j = 0 To width
    For i = 0 To length
        returnArray(i, j) = arrs(j)(i)
    Next i

Next j
connectArrays = returnArray
End Function

'Matrix has to be first variable in arrs (logic is the same as in ceonnextMatrix()) + you can determine how many columns need to be left out of the arrays with leaveStartColOutArray
Function connectMatrixWithArray(arrs As Variant, Optional leaveStartColOutArray As Variant = 0) As Variant
Dim returnArray As Variant
Dim width As Integer
Dim widCounter As Integer
Dim leaveStartColOut As Integer
length = UBound(arrs(0), 1)
width = UBound(arrs(0), 2)
For i = LBound(arrs) + 1 To UBound(arrs)
    width = width + UBound(arrs(i), 1) + 1
Next i

ReDim returnArray(0 To length, 0 To width - 1)

For i = LBound(arrs) To UBound(arrs)
    If i = LBound(arrs) Then
        For j = 0 To length
            For k = 0 To UBound(arrs(i), 2)
                returnArray(j, widCounter + k) = arrs(i)(j, k)
            Next k
        Next j
        widCounter = widCounter + k
    Else
        If Not IsEmpty(leaveStartColOutArray) Then leaveStartColOut = leaveStartColOutArray(i - 1)
        For j = 0 To UBound(arrs(i)) - leaveStartColOut
            returnArray(0, widCounter + j) = arrs(i)(j + leaveStartColOut)
        Next j
        widCounter = widCounter + j
    End If
Next i

connectMatrixWithArray = returnArray
End Function
    'higher complexity operations +++++++++++++++++++++++++++++++++++++++++++++
'returns one array of the appropiate values for each different value in the orderColumn
Function splitArray(arr As Variant, byOrderCol As Integer) As Variant

Dim pivotArray() As Variant
Dim returnArray As Variant

length = UBound(arr, 1)
lLength = LBound(arr, 1)
counter = lLength
For i = lLength To length
    index = inArray(pivotArray, arr(i, byOrderCol))
    If index = -1 Then
        ReDim Preserve pivotArray(lLength To counter)
        pivotArray(counter) = arr(i, byOrderCol)               'infers orderValue
        counter = counter + 1
    End If
Next i

newLength = UBound(pivotArray)
ReDim returnArray(lLength To newLength)
For i = lLength To newLength
    returnArray(i) = filterArray(arr, "argPivot", Array(byOrderCol, pivotArray(i)))
Next i
splitArray = returnArray
End Function

Function uniqueInArray(arr As Variant) As Variant
Dim returnArray() As Variant
Dim counter As Integer

length = UBound(arr)

For i = 0 To length
    If inArray(returnArray, arr(i)) = -1 Then
        ReDim Preserve returnArray(0 To counter)
        returnArray(counter) = arr(i)
        counter = counter + 1
    End If
Next i

uniqueInArray = returnArray
End Function

Function uniqueInMatrix(mat As Variant) As Variant
    Dim length As Integer, width As Integer
    Dim uniqueArrays As Variant
    Dim sizes() As Integer
    Dim maxSize As Integer
    Dim returnArray As Variant
    
    length = UBound(mat, 1)
    width = UBound(mat, 2)
    
    ReDim uniqueArrays(0 To width)
    ReDim sizes(0 To width)
    For i = 0 To width
        uniqueArrays(i) = uniqueInArray(getColumnFromMatrix(mat, i))
        sizes(i) = UBound(uniqueArrays(i))
        If maxSize <= sizes(i) Then maxSize = sizes(i)
    Next i
    
    ReDim returnArray(0 To Size, 0 To width)
    
    For i = 0 To width
        For j = 0 To sizes(i)
            returnArray(j, i) = uniqueArrays(i)(j)
        Next j
    Next i
    
    uniqueInMatrix = returnArray
End Function

Function doItemsInArrayMatch(arr1 As Variant, arr2 As Variant) As Boolean
    Dim i As Integer
    For i = 0 To UBound(arr1)
        If inArray(arr2, arr1(i)) = -1 Then
            doItemsInArrayMatch = False
            Exit Function
        End If
    Next i
    
    For i = 0 To UBound(arr2)
        If inArray(arr1, arr2(i)) = -1 Then
            doItemsInArrayMatch = False
            Exit Function
        End If
    Next i
    
    doItemsInArrayMatch = True
End Function

Function sizesOfUniqueArraysInMat(mat As Variant)
    Dim width As Integer
    Dim uniqueArrays As Variant
    Dim sizes() As Integer
    
    width = UBound(mat, 2)
    
    ReDim uniqueArrays(0 To width)
    ReDim sizes(0 To width)
    For i = 0 To width
        uniqueArrays(i) = uniqueInArray(getColumnFromMatrix(mat, i))
        sizes(i) = UBound(uniqueArrays(i))
    Next i
    
    sizesOfUniqueArraysInMat = sizes
End Function


'returns the index of the first element equal the arg, if arg not in arr then return -1
Function inArray(arr As Variant, arg As Variant) As Integer
On Error GoTo err

For Each i In arr
    If i = arg Then
        inArray = counter
        Exit Function
    End If
    counter = counter + 1
Next i
err:
On Error GoTo 0
inArray = -1
End Function

'Conditions: create Function returning bool value if criteria is met and use name of this function as argument:=filterfunction
'e.g:
Function argPivot(argArr As Variant, filterCriteria As Variant) As Boolean
If argArr(filterCriteria(0)) = filterCriteria(1) Then
    argPivot = True
Else
    argPivot = False
End If
End Function
Function filterArray(arr As Variant, filterFunction As String, filterCriteria1 As Variant) As Variant
Dim helpArray As Variant
Dim criteriaArray As Variant

Dim lowlength As Integer
Dim lowwidth As Integer
Dim uplength As Integer
Dim upwidth As Integer
Dim dimCounter As Integer

lowlength = LBound(arr, 1)
lowwidth = LBound(arr, 1)
uplength = UBound(arr, 1)
upwidth = UBound(arr, 2)

ReDim criteriaArray(lowlength To upwidth)
dimCounter = lowlength
For i = lowlength To uplength
    For j = lowwidth To upwidth
        criteriaArray(j) = arr(i, j)
    Next j
    If Application.Run(filterFunction, criteriaArray, filterCriteria1) Then
        dimCounter = dimCounter + 1
    End If
Next i

If dimCounter > lowlength Then
    dimCounter = dimCounter - 1 'one gets added too much after last satisfied criteria
Else    'no value matched criteria
    Exit Function
End If

ReDim helpArray(lowlength To dimCounter, lowwidth To upwidth)
dimCounter = lowlength
For i = lowlength To uplength
    For j = lowwidth To upwidth
        criteriaArray(j) = arr(i, j)
    Next j
    If Application.Run(filterFunction, criteriaArray, filterCriteria1) Then
        For j = lowwidth To upwidth
            helpArray(dimCounter, j) = arr(i, j)
        Next j
        dimCounter = dimCounter + 1
    End If
Next i

filterArray = helpArray
End Function

'can sort 1 to 2 dimensional arrays
Sub sortArray(ByRef oldArray As Variant, Optional sortColumn As Integer = 0)

Dim length As Integer
Dim width As Integer
Dim helpArray As Variant

Dim i As Integer

length = UBound(oldArray, 1)

On Error Resume Next
width = UBound(oldArray, 2)

If width = 0 Then
    On Error GoTo 0
    ReDim helpArray(0 To length, 0)
    For i = LBound(oldArray) To length
        helpArray(i, 0) = oldArray(i)
    Next i
    'NoramlSort helpArray, sortColumn
    QuickSort helpArray, 0, length, sortColumn
    For i = LBound(oldArray) To length
        oldArray(i) = helpArray(i, 0)
    Next i
    Exit Sub
Else
    On Error GoTo 0
    'NoramlSort oldArray, sortColumn
    QuickSort oldArray, 0, length, sortColumn
End If

End Sub

Sub swapIndex(ByRef oldArray As Variant, left As Integer, right As Integer)
    Dim width As Integer
    Dim helpArray As Variant
    
    Dim j As Integer
    
    width = UBound(oldArray, 2)
    
    ReDim helpArray(0 To width)
    
    For j = 0 To width
        helpArray(j) = oldArray(left, j)
        oldArray(left, j) = oldArray(right, j)
        oldArray(right, j) = helpArray(j)
    Next j
End Sub

Sub QuickSort(ByRef arr, left As Integer, right As Integer, sortColumn As Integer)
  Dim varPivot As Variant
  Dim leftPart As Integer
  Dim rightPart As Integer
  Dim formula As Integer
  
  leftPart = left
  rightPart = right
  formula = (left + right) / 2
  varPivot = arr(formula, sortColumn)
  
  Do While leftPart <= rightPart
    Do While arr(leftPart, sortColumn) < varPivot And leftPart < right
      leftPart = leftPart + 1
    Loop
    Do While varPivot < arr(rightPart, sortColumn) And rightPart > left
      rightPart = rightPart - 1
    Loop
    If leftPart <= rightPart Then
      swapIndex arr, leftPart, rightPart
      leftPart = leftPart + 1
      rightPart = rightPart - 1
    End If
  Loop
  If left < rightPart Then QuickSort arr, left, rightPart, sortColumn
  If leftPart < right Then QuickSort arr, leftPart, right, sortColumn
End Sub

Sub NoramlSort(ByRef arr, sortColumn As Integer)

Dim length As Integer

Dim pivot As Variant
Dim oldPivot As Variant

Dim right As Integer

length = UBound(arr, 1)

For i = 0 To length
    pivot = arr(i, sortColumn)
    For j = i + 1 To length
        If arr(j, sortColumn) >= oldPivot And arr(j, sortColumn) < pivot Then
            pivot = arr(j, sortColumn)
            right = j
        End If
    Next j
    
    oldPivot = pivot
    If arr(i, sortColumn) > pivot Then swapIndex arr, i, right
Next i
End Sub

''------------------------------------------------------------------------------------------------------------------------------------
'RANDOM OPERATIONS

Function returnRandomNotNullCol(arr As Variant, row As Integer, width As Integer, Optional startCol As Integer = 0) As Integer
rndNumber = (-1) ^ Int(2 * Rnd + 1)
rndCol = Int((width - startCol + 1) * Rnd + startCol)
j = rndCol
Do While j >= startCol And j <= width
    If check(arr(row, j), "Double", "", 0, "") > 0 Then
        returnRandomNotNullCol = j
        Exit Function
    End If
    j = j + rndNumber
Loop
j = rndCol - rndNumber
Do While j >= startCol And j <= width
    If check(arr(row, j), "Double", "", 0, "") > 0 Then
        returnRandomNotNullCol = j
        Exit Function
    End If
    j = j - rndNumber
Loop
returnRandomNotNullCol = -1
End Function








