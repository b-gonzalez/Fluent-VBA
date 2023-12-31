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
Private WithEvents pTests As cTests
Attribute pTests.VB_VarHelpID = -1
Private WithEvents pTestOf As cTests
Attribute pTestOf.VB_VarHelpID = -1
Private WithEvents pTestOfResults As cTests
Attribute pTestOfResults.VB_VarHelpID = -1
Private TestCount As Long
Private testCountOf As Long
Private i As Long

Private Sub Class_Initialize()
    TestCount = 0
    testCountOf = 0
End Sub

Public Property Set setFluent(value As cFluent)
    Set pFluentEvent = value
    Set pTests = pFluentEvent.Meta.Tests
End Property

Public Property Set setFluentOf(value As cFluentOf)
    Set pFluentEventOf = value
    Set pTestOf = pFluentEventOf.Meta.Tests
End Property

Public Property Set setFluentEventOfResult(value As cFluentOf)
    Set pFluentEventOfResult = value
    Set pTestOfResults = pFluentEventOfResult.Meta.Tests
End Property

Private Sub pTestOf_TestPassed(ByVal test As cTest)
    Debug.Assert pFluentEventOf.Meta.Tests.Result = test.Result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestFailed(ByVal test As cTest)
    Debug.Assert pFluentEventOf.Meta.Tests.Result = test.Result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestCompleted(ByVal test As cTest)
    testCountOf = testCountOf + 1
    Debug.Assert pFluentEventOf.Meta.Tests.Count = testCountOf
End Sub

Private Sub pTest_TestPassed(ByVal test As cTest)
    Debug.Assert pFluentEvent.Meta.Tests.Result = test.Result
    Debug.Assert test.ExpectedValue = test.ActualValue
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestFailed(ByVal test As cTest)
    Debug.Assert pFluentEvent.Meta.Tests.Result = test.Result
    Debug.Assert test.ExpectedValue <> test.ActualValue
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestCompleted(ByVal test As cTest)
    TestCount = TestCount + 1
    Debug.Assert pFluentEvent.Meta.Tests.Count = TestCount
End Sub

Private Sub pTestOfResult_TestPassed(ByVal test As cTest)
    Debug.Assert test.ExpectedValue = test.ActualValue
End Sub

Private Sub pTestOfResult_TestFailed(ByVal test As cTest)
    Debug.Assert test.ExpectedValue <> test.ActualValue
End Sub
