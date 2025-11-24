VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let testingValue(ByVal value As Variant)
End Property

Public Property Set testingValue(ByVal value As Variant)
End Property

Public Property Let testingInput(ByVal value As Variant)
End Property

Public Property Set testingInput(ByVal value As Variant)
End Property

Public Property Let lowerVal(ByVal value As Variant)
End Property

Public Property Let higherVal(ByVal value As Variant)
End Property

Public Property Let FluentPath(ByVal value As String)
End Property

Public Property Let result(ByVal value As Variant)
End Property

Public Property Let functionName(ByVal value As String)
End Property

Public Property Let strTestValue(ByVal value As String)
End Property

Public Property Let StrTestInput(ByVal value As String)
End Property

Public Property Let strTestValuePretty(ByVal value As String)
End Property

Public Property Let StrTestInputPretty(ByVal value As String)
End Property

Public Property Let HasNull(ByVal value As Boolean)
End Property

Public Property Let Category(ByVal value As String)
End Property

Public Property Let TestName(ByVal value As String)
End Property

Public Property Let TestValueSet(ByVal value As Boolean)
End Property

Public Property Let HasEmpty(ByVal value As Boolean)
End Property

Public Property Let cleanedTestValue(ByVal value As String)
End Property

Public Property Let cleanedTestInput(ByVal value As String)
End Property

Public Property Let HasCleanTestValue(ByVal value As Boolean)
End Property

Public Property Let HasCleanTestInput(ByVal value As Boolean)
End Property

Public Property Let HasCleanTestStrings(ByVal value As Boolean)
End Property

Public Property Let Actual(ByVal value As String)
End Property

Public Property Let HasSelfReferential(ByVal value As Variant)
End Property

Public Property Let TestingValueIsSelfReferential(ByVal value As Variant)
End Property

Public Property Let TestingInputIsSelfReferential(ByVal value As Variant)
End Property
