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

Public Property Get testingValue() As Variant
End Property

Public Property Get testingInput() As Variant
End Property

Public Property Get lowerVal() As Variant
End Property

Public Property Get higherVal() As Variant
End Property

Public Property Get FluentPath() As String
End Property

Public Property Get result() As Variant
End Property

Public Property Get functionName() As String
End Property

Public Property Get strTestValue() As String
End Property

Public Property Get StrTestInput() As String
End Property

Public Property Get strTestValuePretty() As String
End Property

Public Property Get StrTestInputPretty() As String
End Property

Public Property Get HasNull() As Boolean
End Property

Public Property Get Category() As String
End Property

Public Property Get TestName() As String
End Property

Public Property Get TestValueSet() As Boolean
End Property

Public Property Get HasEmpty() As Boolean
End Property

Public Property Get cleanedTestValue() As String
End Property

Public Property Get cleanedTestInput() As String
End Property

Public Property Get HasCleanTestValue() As Boolean
End Property

Public Property Get HasCleanTestInput() As Boolean
End Property

Public Property Get HasCleanTestStrings() As Boolean
End Property

Public Property Get Actual() As String
End Property

Public Property Get HasSelfReferential() As Variant
End Property

Public Property Get TestingValueIsSelfReferential() As Variant
End Property

Public Property Get TestingInputIsSelfReferential() As Variant
End Property
