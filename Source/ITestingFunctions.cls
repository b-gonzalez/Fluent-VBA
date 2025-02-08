VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestingFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum flAlgorithm
    flRecursive = 2 ^ 0
    flIterative = 2 ^ 1
End Enum

Public Enum flTestType
    flIsArray = 2 ^ 0
    flIsNotArray = 2 ^ 1
    flIsDataStructure = 2 ^ 2
    flIsErrObject = 2 ^ 3
    flIsNotErrObject = 2 ^ 4
    flIsNotDataStructure = 2 ^ 5
    flIsNull = 2 ^ 6
    flIsNotNull = 2 ^ 7
    flIsNumeric = 2 ^ 8
    flIsNotNumeric = 2 ^ 9
    flIsObject = 2 ^ 10
    flIsNotObject = 2 ^ 11
    flIsSelfReferential = 2 ^ 12
    flIsNotSelfReferential = 2 ^ 13
    flIsString = 2 ^ 14
    flIsNotString = 2 ^ 15
    flIsVariant = 2 ^ 16
    flIsNotVariant = 2 ^ 17
End Enum
Public Property Let TestValue(ByVal value As Variant)
End Property

Public Property Let TestValueClean(ByVal value As Variant)
End Property

Public Property Set TestValue(ByVal value As Variant)
End Property

Public Property Get TestValue() As Variant
End Property

Public Property Let TestValueSet(ByVal value As Boolean)
End Property

Public Function SetDefaultFuncVal() As Variant
End Function

Public Property Set setMeta(ByVal value As cMeta)
End Property

Public Property Set Expressions(ByVal value As IExpressions)
End Property

Public Function GetCleanedString(ByVal value As Variant) As String
End Function

Public Function GetTempTestValue(ByVal testingValue As Variant) As Variant
End Function

Public Function GetTempTestInput(ByVal testingInput As Variant) As Variant
End Function

Private Function inputIsValid(ByVal inputValue As Variant, ByVal inputConditions As flTestType) As Boolean

End Function
