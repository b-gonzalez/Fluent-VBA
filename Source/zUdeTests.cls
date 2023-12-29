VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zUdeTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class module is used to test user-defined events

Option Explicit

Private pFluentEvent As cFluent
Private pFluentEventOf As cFluentOf
Private pFluentEventOfResult As cFluentOf
Private WithEvents pTestResults As cTestResults
Attribute pTestResults.VB_VarHelpID = -1
Private WithEvents pTestResultOf As cTestResults
Attribute pTestResultOf.VB_VarHelpID = -1
Private WithEvents pTestResultOfResults As cTestResults
Attribute pTestResultOfResults.VB_VarHelpID = -1
Private TestCount As Long
Private testCountOf As Long
Private i As Long

Private Sub Class_Initialize()
    TestCount = 0
    testCountOf = 0
End Sub

Public Property Set setFluent(value As cFluent)
    Set pFluentEvent = value
    Set pTestResults = pFluentEvent.Meta.TestResults
End Property

Public Property Set setFluentOf(value As cFluentOf)
    Set pFluentEventOf = value
    Set pTestResultOf = pFluentEventOf.Meta.TestResults
End Property

Public Property Set setFluentEventOfResult(value As cFluentOf)
    Set pFluentEventOfResult = value
    Set pTestResultOfResults = pFluentEventOfResult.Meta.TestResults
End Property

Private Sub pTestResultOf_TestPassed(ByVal TestResult As cTestResult)
    Debug.Assert pFluentEventOf.Meta.TestResults.Result = TestResult.Result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = TestResult.FluentPath
End Sub

Private Sub pTestResultOf_TestFailed(ByVal TestResult As cTestResult)
    Debug.Assert pFluentEventOf.Meta.TestResults.Result = TestResult.Result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = TestResult.FluentPath
End Sub

Private Sub pTestResultOf_TestCompleted(ByVal TestResult As cTestResult)
    testCountOf = testCountOf + 1
    Debug.Assert pFluentEventOf.Meta.TestCount = testCountOf
End Sub

Private Sub pTestResult_TestPassed(ByVal TestResult As cTestResult)
    Debug.Assert pFluentEvent.Meta.TestResults.Result = TestResult.Result
    Debug.Assert TestResult.ExpectedValue = TestResult.ActualValue
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = TestResult.FluentPath
End Sub

Private Sub pTestResult_TestFailed(ByVal TestResult As cTestResult)
    Debug.Assert pFluentEvent.Meta.TestResults.Result = TestResult.Result
    Debug.Assert TestResult.ExpectedValue <> TestResult.ActualValue
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = TestResult.FluentPath
End Sub

Private Sub pTestResult_TestCompleted(ByVal TestResult As cTestResult)
    TestCount = TestCount + 1
    Debug.Assert pFluentEvent.Meta.TestCount = TestCount
End Sub

Private Sub pTestResultOfResult_TestPassed(ByVal TestResult As cTestResult)
    Debug.Assert TestResult.ExpectedValue = TestResult.ActualValue
End Sub

Private Sub pTestResultOfResult_TestFailed(ByVal TestResult As cTestResult)
    Debug.Assert TestResult.ExpectedValue <> TestResult.ActualValue
End Sub
