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
Private pCleanedTestValue As String
Private pCleanedTestInput As String
Private phasCleanTestValue As Boolean
Private phasCleanTestInput As Boolean
Private phasCleanTestStrings As Boolean
Private pAcutal As String

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

Private Property Let ITest_CleanedTestValue(value As String)
    pCleanedTestValue = value
End Property

Private Property Get ITest_CleanedTestValue() As String
    ITest_CleanedTestValue = pCleanedTestValue
End Property

Private Property Let ITest_CleanedTestInput(value As String)
    pCleanedTestInput = value
End Property

Private Property Get ITest_CleanedTestInput() As String
    ITest_CleanedTestInput = pCleanedTestInput
End Property

Private Property Let ITest_HasCleanTestValue(ByVal value As Boolean)
    phasCleanTestValue = value
End Property

Private Property Get ITest_HasCleanTestValue() As Boolean
    ITest_HasCleanTestValue = phasCleanTestValue
End Property

Private Property Let ITest_HasCleanTestInput(ByVal value As Boolean)
    phasCleanTestInput = False
End Property

Private Property Get ITest_HasCleanTestInput() As Boolean
    ITest_HasCleanTestInput = phasCleanTestInput
End Property

Private Property Let ITest_HasCleanTestStrings(ByVal value As Boolean)
    phasCleanTestStrings = value
End Property

Private Property Get ITest_HasCleanTestStrings() As Boolean
    ITest_HasCleanTestStrings = phasCleanTestStrings
End Property

Private Property Let ITest_Actual(value As String)
    pAcutal = value
End Property

Private Property Get ITest_Actual() As String
    ITest_Actual = pAcutal
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

Private Property Let ITestDev_TestValueRecur(value As String)
    pTestValueRecur = value
End Property

Private Property Get ITestDev_TestValueRecur() As String
    ITestDev_TestValueRecur = pTestValueRecur
End Property

Private Property Let ITestDev_TestValueIter(value As String)
    pTestValueIter = value
End Property

Private Property Get ITestDev_TestValueIter() As String
    ITestDev_TestValueIter = pTestValueIter
End Property

Private Property Let ITestDev_TestInputRecur(value As String)
    pTestInputRecur = value
End Property

Private Property Get ITestDev_TestInputRecur() As String
    ITestDev_TestInputRecur = pTestInputRecur
End Property

Private Property Let ITestDev_TestInputIter(value As String)
    pTestInputIter = value
End Property

Private Property Get ITestDev_TestInputIter() As String
    ITestDev_TestInputIter = pTestInputIter
End Property

'PUBLIC PROPERTIES

Public Property Let testingValue(ByVal value As Variant)
    ITest_testingValue = value
End Property

Public Property Set testingValue(ByVal value As Variant)
    Set ITest_testingValue = value
End Property

Public Property Get testingValue() As Variant
    If TypeOf pTestingValue Is Object  Then
        Set testingValue = ITest_testingValue
    Else
        testingValue = ITest_testingValue
    End If
End Property

Public Property Let testingInput(ByVal value As Variant)
    ITest_testingInput = value
End Property

Public Property Set testingInput(ByVal value As Variant)
    Set ITest_testingInput = value
End Property

Public Property Get testingInput() As Variant
    If TypeOf pTestingInput Is Object  Then
        Set testingInput = ITest_testingInput
    Else
        testingInput = ITest_testingInput
    End If
End Property

Public Property Let lowerVal(ByVal value As Double)
    ITest_LowerVal = value
End Property

Public Property Get lowerVal() As Double
    lowerVal = ITest_LowerVal
End Property

Public Property Let higherVal(ByVal value As Double)
    ITest_HigherVal = value
End Property

Public Property Get higherVal() As Double
    higherVal = ITest_HigherVal
End Property

Public Property Let FluentPath(ByVal value As String)
    ITest_FluentPath = value
End Property

Public Property Get FluentPath() As String
    FluentPath = ITest_FluentPath
End Property

Public Property Let result(ByVal value As Variant)
    ITest_result = value
End Property

Public Property Get result() As Variant
    result = ITest_result
End Property

Public Property Let functionName(ByVal value As String)
    ITest_functionName = value
End Property

Public Property Get functionName() As String
    functionName = ITest_functionName
End Property

Public Property Let strTestValue(ByVal value As String)
    ITest_StrTestValue = value
End Property

Public Property Get strTestValue() As String
    strTestValue = ITest_StrTestValue
End Property

Public Property Let StrTestInput(ByVal value As String)
    ITest_StrTestInput = value
End Property

Public Property Get StrTestInput() As String
    StrTestInput = ITest_StrTestInput
End Property

Public Property Let HasNull(ByVal value As Boolean)
    ITest_HasNull = value
End Property

Public Property Get HasNull() As Boolean
    HasNull = ITest_HasNull
End Property

Public Property Let Category(ByVal value As String)
    ITest_Category = value
End Property

Public Property Get Category() As String
    Category = ITest_Category
End Property

Public Property Let TestName(ByVal value As String)
    ITest_TestName = value
End Property

Public Property Get TestName() As String
    TestName = ITest_TestName
End Property

Public Property Let TestValueSet(ByVal value As Boolean)
    ITest_TestValueSet = value
End Property

Public Property Get TestValueSet() As Boolean
    TestValueSet = ITest_TestValueSet
End Property

Public Property Let HasEmpty(ByVal value As Boolean)
    ITest_HasEmpty = value
End Property

Public Property Get HasEmpty() As Boolean
    HasEmpty = ITest_HasEmpty
End Property

Public Property Let cleanedTestValue(value As String)
    ITest_CleanedTestValue = value
End Property

Public Property Get cleanedTestValue() As String
    cleanedTestValue = ITest_CleanedTestValue
End Property

Public Property Let cleanedTestInput(value As String)
    ITest_CleanedTestInput = value
End Property

Public Property Get cleanedTestInput() As String
    cleanedTestInput = ITest_CleanedTestInput
End Property

Public Property Let HasCleanTestValue(ByVal value As Boolean)
    ITest_HasCleanTestValue = value
End Property

Public Property Get HasCleanTestValue() As Boolean
    HasCleanTestValue = ITest_HasCleanTestValue
End Property

Public Property Let HasCleanTestInput(ByVal value As Boolean)
    ITest_HasCleanTestInput = value
End Property

Public Property Get HasCleanTestInput() As Boolean
    HasCleanTestInput = ITest_HasCleanTestInput
End Property

Public Property Let HasCleanTestStrings(ByVal value As Boolean)
    ITest_HasCleanTestStrings = value
End Property

Public Property Get HasCleanTestStrings() As Boolean
    HasCleanTestStrings = ITest_HasCleanTestStrings
End Property

Public Property Let Actual(value As String)
    ITest_Actual = value
End Property

Public Property Get Actual() As String
    Actual = ITest_Actual
End Property
