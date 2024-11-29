VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let testingValue(ByVal value As Variant)
End Property

Public Property Set testingValue(ByVal value As Variant)
End Property

Public Property Get testingValue() As Variant
End Property

Public Property Let testingInput(ByVal value As Variant)
End Property

Public Property Set testingInput(ByVal value As Variant)
End Property

Public Property Get testingInput() As Variant
End Property

Public Property Let lowerVal(ByVal value As Variant)
End Property

Public Property Get lowerVal() As Variant
End Property

Public Property Let higherVal(ByVal value As Variant)
End Property

Public Property Get higherVal() As Variant
End Property

Public Property Let FluentPath(ByVal value As String)
End Property

Public Property Get FluentPath() As String
End Property

Public Property Let result(ByVal value As Variant)
End Property

Public Property Get result() As Variant
End Property

Public Property Let functionName(ByVal value As String)
End Property

Public Property Get functionName() As String
End Property

Public Property Let strTestValue(ByVal value As String)
End Property

Public Property Get strTestValue() As String
End Property

Public Property Let StrTestInput(ByVal value As String)
End Property

Public Property Get StrTestInput() As String
End Property

Public Property Let HasNull(ByVal value As Boolean)
End Property

Public Property Get HasNull() As Boolean
End Property

Public Property Let Category(ByVal value As String)
End Property

Public Property Get Category() As String
End Property

Public Property Let TestName(ByVal value As String)
End Property

Public Property Get TestName() As String
End Property

Public Property Let TestValueSet(ByVal value As Boolean)
End Property

Public Property Get TestValueSet() As Boolean
End Property

Public Property Let HasEmpty(ByVal value As Boolean)
End Property

Public Property Get HasEmpty() As Boolean
End Property

Public Property Let cleanedTestValue(ByVal value As String)
End Property

Public Property Get cleanedTestValue() As String
End Property

Public Property Let cleanedTestInput(ByVal value As String)
End Property

Public Property Get cleanedTestInput() As String
End Property

Public Property Let HasCleanTestValue(ByVal value As Boolean)
End Property

Public Property Get HasCleanTestValue() As Boolean
End Property

Public Property Let HasCleanTestInput(ByVal value As Boolean)
End Property

Public Property Get HasCleanTestInput() As Boolean
End Property

Public Property Let HasCleanTestStrings(ByVal value As Boolean)
End Property

Public Property Get HasCleanTestStrings() As Boolean
End Property

Public Property Let Actual(ByVal value As String)
End Property

Public Property Get Actual() As String
End Property
