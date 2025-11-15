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
