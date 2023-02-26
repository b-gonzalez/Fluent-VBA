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

'NOTE: pTestResult implements a check to see that the
'expected and actual values are equal to or not equal
'to each other when a test passes or fails respectively.
'This check is not implemented in the testResultOf events
'This is because this is not possible to do as the tests
'are currently written. The tests as written in mTests use
'cFluentOf objects for the initial tests. The result of those
'initial tests are then valided by cFluent objects. So cFluent
'objects have access to cFluentOf objects tests results when
'their event is raised, but not vice versa.

Option Explicit

Private pFluentEvent As cFluent
Private pFluentEventOf As cFluentOf
Private pFluentEventOfResult As cFluentOf
Private WithEvents pTestResult As cTestResult
Private WithEvents pTestResultOf As cTestResult
Private WithEvents pTestResultOfResult As cTestResult
Attribute pTestResultOfResult.VB_VarHelpID = -1
Private testCount As Long
Private testCountOf As Long
Private i As Long

Private Sub Class_Initialize()
    testCount = 0
    testCountOf = 0
End Sub

Public Property Set setFluent(value As cFluent)
    Set pFluentEvent = value
    Set pTestResult = pFluentEvent.Meta.TestResult
End Property

Public Property Set setFluentOf(value As cFluentOf)
    Set pFluentEventOf = value
    Set pTestResultOf = pFluentEventOf.Meta.TestResult
End Property

Public Property Set setFluentEventOfResult(value As cFluentOf)
    Set pFluentEventOfResult = value
    Set pTestResultOfResult = pFluentEventOfResult.Meta.TestResult
End Property

Private Sub pTestResultOf_TestPassed(ByVal TestDict As Scripting.Dictionary)
    Debug.Assert pFluentEventOf.Meta.TestResult.Result = TestDict("result")
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = TestDict("fluentPath")
End Sub

Private Sub pTestResultOf_TestFailed(ByVal TestDict As Scripting.Dictionary)
    Debug.Assert pFluentEventOf.Meta.TestResult.Result = TestDict("result")
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = TestDict("fluentPath")
End Sub

Private Sub pTestResultOf_TestCompleted(ByVal TestDict As Scripting.Dictionary)
    testCountOf = testCountOf + 1
    Debug.Assert pFluentEventOf.Meta.testCount = testCountOf
End Sub

Private Sub pTestResult_TestPassed(ByVal TestDict As Scripting.Dictionary)
    Debug.Assert pFluentEvent.Meta.TestResult.Result = TestDict("result")
    Debug.Assert TestDict("expected") = TestDict("actual")
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = TestDict("fluentPath")
End Sub

Private Sub pTestResult_TestFailed(ByVal TestDict As Scripting.Dictionary)
    Debug.Assert pFluentEvent.Meta.TestResult.Result = TestDict("result")
    Debug.Assert TestDict("expected") <> TestDict("actual")
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = TestDict("fluentPath")
End Sub

Private Sub pTestResult_TestCompleted(ByVal TestDict As Scripting.Dictionary)
    testCount = testCount + 1
    Debug.Assert pFluentEvent.Meta.testCount = testCount
End Sub

Private Sub pTestResultOfResult_TestPassed(ByVal TestDict As Scripting.Dictionary)
    Debug.Assert TestDict("expected") = TestDict("actual")
End Sub

Private Sub pTestResultOfResult_TestFailed(ByVal TestDict As Scripting.Dictionary)
    Debug.Assert TestDict("expected") <> TestDict("actual")
End Sub
