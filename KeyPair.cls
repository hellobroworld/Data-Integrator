VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pBaseTableName As String
Private pBaseKeyColumnIndex As Integer
Private pBaseKeyColumnName As String

Private pCompareTableNames() As String
Private pCompareKeyColumnIndices() As Integer
Private pCompareKeyNames() As String

Property Get baseTableName() As String
    baseTableName = pBaseTableName
End Property

Property Get baseKeyColumnIndex() As String
    baseKeyColumnIndex = pBaseKeyColumnIndex
End Property

Property Get baseKeyColumnName() As String
    baseKeyColumnName = pBaseKeyColumnName
End Property


Property Get compareTableNames() As String()
    compareTableNames = pCompareTableNames
End Property

Property Get compareKeyColumnIndices() As Integer()
    compareKeyColumnIndices = pCompareKeyColumnIndices
End Property
Property Get compareKeyNames() As String()
    compareKeyNames = pCompareKeyNames
End Property


Sub init(baseName As String, baseKeyIndex As Integer, baseKeyName As String, compareNames() As String, compareKeyIndeces() As Integer, compareKeyNames() As String)
    pBaseTableName = baseName
    pBaseKeyColumnIndex = baseKeyIndex
    pBaseKeyColumnName = baseKeyName
    
    pCompareTableNames = compareNames
    pCompareKeyColumnIndices = compareKeyIndeces
    pCompareKeyNames = compareKeyNames
End Sub

