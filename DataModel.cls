VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pKeyPair As keyPair
Private pAttributesPair As Variant        '3 Dim Array (1: Attributes, 2: Tables, 3: {Index of Attribute, ColumnName}

Property Get keyPair() As keyPair
    Set keyPair = pKeyPair
End Property

Property Get attributePairs() As Variant
    attributePairs = pAttributesPair
End Property

Sub init(diMasks() As DIMaskModel)   'diMaskData: Array of matrices (matrices are the tables from DI Mask Worksheet)

    Set pKeyPair = initKeyPair(diMasks)
    pAttributesPair = initAttributePair(diMasks)
    
End Sub

Private Function initKeyPair(diMasks() As DIMaskModel) As keyPair
    Dim keyIndex As Integer
    Dim colIndex As Integer
    Dim length As Integer
    Dim i, j As Integer
    
    Dim diMask As DIMaskModel
    Dim tableArray As Variant
    Dim lengthTable As Integer
    
    Dim baseTableName As String
    Dim baseKeyIndex As Integer
    Dim baseKeyName As String

    Dim compareTableNames() As String
    Dim compareKeyIndeces() As Integer
    Dim compareKeyNames() As String
    
    Dim keyPair As New keyPair
    
    length = UBound(diMasks)
    
    ReDim compareTableNames(1 To length)
    ReDim compareKeyIndeces(1 To length)
    ReDim compareKeyNames(1 To length)
    
    For i = 0 To length
        Set diMask = diMasks(i)
        tableArray = diMask.dIMaskData
        lengthTable = UBound(tableArray)
        For j = 0 To lengthTable
            If Not IsEmpty(tableArray(j, diMask.keyCol)) Then
                If i = 0 Then
                    baseTableName = diMask.tableName
                    baseKeyIndex = j
                    baseKeyName = tableArray(j, diMask.columnCol)
                Else
                    compareTableNames(i) = diMask.tableName
                    compareKeyIndeces(i) = j
                    compareKeyNames(i) = tableArray(j, diMask.columnCol)
                End If
                j = lengthTable
            End If
        Next j
    Next i
    
    keyPair.init baseTableName, baseKeyIndex, baseKeyName, compareTableNames, compareKeyIndeces, compareKeyNames
    Set initKeyPair = keyPair
End Function

Private Function initAttributePair(diMasks() As DIMaskModel) As Variant   'return: (Column0 to Column1 to ColumnN, ColumnIndex0 to ColumnIndex1 to ColumnIndexN)
    Dim diMask As DIMaskModel
    Dim returnArray As Variant
    
    Dim attributes As Variant
    Dim baseAttributeColumn As Variant
    Dim baseColumns As Variant
    Dim baseAttribute As String
    
    Dim otherAttributeColumn As Variant
    Dim otherColumns As Variant
    
    Dim numberOfTables As Integer
    Dim numberOfAttributes As Integer

    Dim indexOfAttribute As Integer
    
    Dim i As Integer, j As Integer
    
    numberOfTables = UBound(diMasks)
    baseAttributeColumn = getColumnFromMatrix(diMasks(0).dIMaskData, diMasks(0).attributeCol)
    attributes = removeNARowsFromArray(baseAttributeColumn)
    If Not IsEmpty(attributes) Then
    
        baseColumns = getColumnFromMatrix(diMasks(0).dIMaskData, diMasks(0).columnCol)
        numberOfAttributes = UBound(attributes)
    
        ReDim returnArray(0 To numberOfAttributes, 0 To numberOfTables, 0 To 1)
        
        For i = 0 To numberOfTables
            If i = 0 Then
                For j = 0 To numberOfAttributes
                    indexOfAttribute = inArray(baseAttributeColumn, attributes(j))
                    returnArray(j, i, 0) = indexOfAttribute
                    returnArray(j, i, 1) = baseColumns(indexOfAttribute)
                Next j
            Else
                otherAttributeColumn = getColumnFromMatrix(diMasks(i).dIMaskData, diMasks(i).attributeCol)
                otherColumns = getColumnFromMatrix(diMasks(i).dIMaskData, diMasks(i).columnCol)
                For j = 0 To numberOfAttributes
                    indexOfAttribute = inArray(otherAttributeColumn, attributes(j))
                    returnArray(j, i, 0) = indexOfAttribute                             'groups columns of different tables in order of given attribute
                    returnArray(j, i, 1) = otherColumns(indexOfAttribute)
                Next j
            End If
        Next i
        
    End If
    initAttributePair = returnArray
    

End Function

