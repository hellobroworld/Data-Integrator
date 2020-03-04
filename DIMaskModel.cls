VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DIMaskModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pColumnCol As Integer
Private pKeyCol As Integer
Private pAttributeCol As Integer
Private pDIMaskData() As Variant        'array with dataBodyRanges of di mask tables

Property Get columnCol() As String
    columnCol = pColumnCol
End Property

Property Get tableName() As String
    tableName = pTableName
End Property

Property Get keyCol() As String
    keyCol = pKeyCol
End Property

Property Get attributeCol() As String
    attributeCol = pAttributeCol
End Property

Property Get dIMaskData() As Variant()
    dIMaskData = pDIMaskData
End Property

Sub init(nTableName As String, nColumnCol As Integer, nKeyCol As Integer, nAttributeCol As Integer, nDIMaskData As Variant)
    pTableName = nTableName
    pColumnCol = nColumnCol
    pKeyCol = nKeyCol
    pAttributeCol = nAttributeCol
    pDIMaskData = nDIMaskData
End Sub
