VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITest

Private pTestingValue As Variant
Private pTestingInput As Variant
Private pFluentPath As String
Private pResult As Variant
Private pFunctionName As String
Private pLowerVal As Double
Private pHigherVal As Double
Private pStrTestValue As String
Private pStrTestInput As String
Private pHasNull As Boolean
Private pNegateValue As Boolean

Private Property Let ITest_testingValue(value As Variant)
    pTestingValue = value
End Property

Private Property Set ITest_testingValue(value As Variant)
    Set pTestingValue = value
End Property

Private Property Get ITest_testingValue() As Variant
    If TypeOf pTestingValue Is Object  Then
        Set ITest_testingValue = pTestingValue
    Else
        ITest_testingValue = pTestingValue
    End If
End Property

Private Property Let ITest_testingInput(value As Variant)
    pTestingInput = value
End Property

Private Property Set ITest_testingInput(value As Variant)
    Set pTestingInput = value
End Property

Private Property Get ITest_testingInput() As Variant
    If TypeOf pTestingInput Is Object  Then
        Set ITest_testingInput = pTestingInput
    Else
        ITest_testingInput = pTestingInput
    End If
End Property

Private Property Let ITest_LowerVal(value As Double)
    pLowerVal = value
End Property

Private Property Get ITest_LowerVal() As Double
    ITest_LowerVal = pLowerVal
End Property

Private Property Let ITest_HigherVal(value As Double)
    pHigherVal = value
End Property

Private Property Get ITest_HigherVal() As Double
    ITest_HigherVal = pHigherVal
End Property

Private Property Let ITest_FluentPath(value As String)
    pFluentPath = value
End Property

Private Property Get ITest_FluentPath() As String
    ITest_FluentPath = pFluentPath
End Property

Private Property Let ITest_result(value As Variant)
    pResult = value
End Property

Private Property Get ITest_result() As Variant
    ITest_result = pResult
End Property

Private Property Let ITest_functionName(value As String)
    pFunctionName = value
End Property

Private Property Get ITest_functionName() As String
    ITest_functionName = pFunctionName
End Property

Private Property Let ITest_StrTestValue(value As String)
    pStrTestValue = value
End Property

Private Property Get ITest_StrTestValue() As String
    ITest_StrTestValue = pStrTestValue
End Property

Private Property Let ITest_StrTestInput(value As String)
    pStrTestInput = value
End Property

Private Property Get ITest_StrTestInput() As String
    ITest_StrTestInput = pStrTestInput
End Property

Private Property Let ITest_HasNull(value As Boolean)
    pHasNull = value
End Property

Private Property Get ITest_HasNull() As Boolean
    ITest_HasNull = pHasNull
End Property

Private Property Let ITest_NegateValue(value As Boolean)
    pNegateValue = value
End Property

Private Property Get ITest_NegateValue() As Boolean
    ITest_NegateValue = pNegateValue
End Property
