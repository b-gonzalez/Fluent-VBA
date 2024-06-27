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
Implements ITestDev

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
Private pCategory As String
Private pTestName As String
Private pResultNumber As Long
Private pTestValueRecur As String
Private pTestValueIter As String
Private pTestInputRecur As String
Private pTestInputIter As String
Private pTestValueSet As Boolean
Private pHasEmpty As Boolean

Private Property Let ITest_testingValue(ByVal value As Variant)
    pTestingValue = value
End Property

Private Property Set ITest_testingValue(ByVal value As Variant)
    Set pTestingValue = value
End Property

Private Property Get ITest_testingValue() As Variant
    If TypeOf pTestingValue Is Object  Then
        Set ITest_testingValue = pTestingValue
    Else
        ITest_testingValue = pTestingValue
    End If
End Property

Private Property Let ITest_testingInput(ByVal value As Variant)
    pTestingInput = value
End Property

Private Property Set ITest_testingInput(ByVal value As Variant)
    Set pTestingInput = value
End Property

Private Property Get ITest_testingInput() As Variant
    If TypeOf pTestingInput Is Object  Then
        Set ITest_testingInput = pTestingInput
    Else
        ITest_testingInput = pTestingInput
    End If
End Property

Private Property Let ITest_LowerVal(ByVal value As Double)
    pLowerVal = value
End Property

Private Property Get ITest_LowerVal() As Double
    ITest_LowerVal = pLowerVal
End Property

Private Property Let ITest_HigherVal(ByVal value As Double)
    pHigherVal = value
End Property

Private Property Get ITest_HigherVal() As Double
    ITest_HigherVal = pHigherVal
End Property

Private Property Let ITest_FluentPath(ByVal value As String)
    pFluentPath = value
End Property

Private Property Get ITest_FluentPath() As String
    ITest_FluentPath = pFluentPath
End Property

Private Property Let ITest_result(ByVal value As Variant)
    pResult = value
End Property

Private Property Get ITest_result() As Variant
    ITest_result = pResult
End Property

Private Property Let ITest_functionName(ByVal value As String)
    pFunctionName = value
End Property

Private Property Get ITest_functionName() As String
    ITest_functionName = pFunctionName
End Property

Private Property Let ITest_StrTestValue(ByVal value As String)
    pStrTestValue = value
End Property

Private Property Get ITest_StrTestValue() As String
    ITest_StrTestValue = pStrTestValue
End Property

Private Property Let ITest_StrTestInput(ByVal value As String)
    pStrTestInput = value
End Property

Private Property Get ITest_StrTestInput() As String
    ITest_StrTestInput = pStrTestInput
End Property

Private Property Let ITest_HasNull(ByVal value As Boolean)
    pHasNull = value
End Property

Private Property Get ITest_HasNull() As Boolean
    ITest_HasNull = pHasNull
End Property

Private Property Let ITest_Category(ByVal value As String)
    pCategory = value
End Property

Private Property Get ITest_Category() As String
    ITest_Category = pCategory
End Property

Private Property Let ITest_TestName(ByVal value As String)
    pTestName = value
End Property

Private Property Get ITest_TestName() As String
    ITest_TestName = pTestName
End Property

Private Property Let ITest_TestValueSet(ByVal value As Boolean)
    pTestValueSet = value
End Property

Private Property Get ITest_TestValueSet() As Boolean
    ITest_TestValueSet = pTestValueSet
End Property

Private Property Let ITest_HasEmpty(ByVal value As Boolean)
    pHasEmpty = value
End Property

Private Property Get ITest_HasEmpty() As Boolean
    ITest_HasEmpty = pHasEmpty
End Property

Private Property Let ITestDev_negateValue(ByVal value As Boolean)
    pNegateValue = value
End Property

Private Property Get ITestDev_negateValue() As Boolean
    ITestDev_negateValue = pNegateValue
End Property

Private Property Let ITestDev_ResultNumber(ByVal value As Long)
    pResultNumber = value
End Property

Private Property Get ITestDev_ResultNumber() As Long
    ITestDev_ResultNumber = pResultNumber
End Property

Public Property Let ITestDev_TestValueRecur(value As String)
    pTestValueRecur = value
End Property

Public Property Get ITestDev_TestValueRecur() As String
    ITestDev_TestValueRecur = pTestValueRecur
End Property

Public Property Let ITestDev_TestValueIter(value As String)
    pTestValueIter = value
End Property

Public Property Get ITestDev_TestValueIter() As String
    ITestDev_TestValueIter = pTestValueIter
End Property

Public Property Let ITestDev_TestInputRecur(value As String)
    pTestInputRecur = value
End Property

Public Property Get ITestDev_TestInputRecur() As String
    ITestDev_TestInputRecur = pTestInputRecur
End Property

Public Property Let ITestDev_TestInputIter(value As String)
    pTestInputIter = value
End Property

Public Property Get ITestDev_TestInputIter() As String
    ITestDev_TestInputIter = pTestInputIter
End Property
