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

Private Const TEST_UNEXPECTED_ERROR As Long = 513

Private WithEvents pTest As cTests
Attribute pTest.VB_VarHelpID = -1
Private WithEvents pTestOf As cTests
Attribute pTestOf.VB_VarHelpID = -1
Private WithEvents pTestOfResults As cTests
Attribute pTestOfResults.VB_VarHelpID = -1

Private pFluentEvent As cFluent
Private pFluentEventOf As cFluentOf
Private pFluentEventOfResult As cFluentOf
Private TestCount As Long
Private TestOfCount As Long
Private TestOfResultsCount As Long
Private i As Long

Private Sub Class_Initialize()
    TestCount = 0
    TestOfCount = 0
    TestOfResultsCount = 0
End Sub

Public Property Set setFluent(value As cFluent)
    Set pFluentEvent = value
    Set pTest = pFluentEvent.Meta.Tests
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
    Debug.Assert pFluentEventOf.Meta.Tests.result = test.result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestFailed(ByVal test As cTest)
    Debug.Assert pFluentEventOf.Meta.Tests.result = test.result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestCompleted(ByVal test As cTest)
    TestOfCount = TestOfCount + 1
    Debug.Assert pFluentEventOf.Meta.Tests.Count = TestOfCount
End Sub

Private Sub pTest_TestPassed(ByVal test As cTest)
    Debug.Assert pFluentEvent.Meta.Tests.result = test.result
'    Debug.Assert test.ExpectedValue = test.ActualValue
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestFailed(ByVal test As cTest)
    Debug.Assert pFluentEvent.Meta.Tests.result = test.result
'    Debug.Assert test.ExpectedValue <> test.ActualValue
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestCompleted(ByVal test As cTest)
    TestCount = TestCount + 1
    Debug.Assert pFluentEvent.Meta.Tests.Count = TestCount
End Sub

Private Sub pTestOfResult_TestPassed(ByVal test As cTest)
'    Debug.Assert test.ExpectedValue = test.ActualValue
End Sub

Private Sub pTestOfResult_TestFailed(ByVal test As cTest)
'    Debug.Assert test.ExpectedValue <> test.ActualValue
End Sub

Private Sub pTestOfResults_TestCompleted(ByVal test As cTest)
    TestOfResultsCount = TestOfResultsCount + 1
End Sub

Private Sub pTestOf_TestUnexpected(ByVal test As cTest)
    'this event should never be raised
'    Err.Raise Number:=TEST_UNEXPECTED_ERROR, Description:="Test unexpected event should never be raised"
    Debug.Assert test.result = "Null" And test.HasNull
End Sub

Private Sub pTestOfResults_TestUnexpected(ByVal test As cTest)
    'this event should never be raised
    Err.Raise Number:=TEST_UNEXPECTED_ERROR, Description:="Test unexpected event should never be raised"
End Sub

Private Sub pTests_TestUnexpected(ByVal test As cTest)
    'this event should never be raised
    Err.Raise Number:=TEST_UNEXPECTED_ERROR, Description:="Test unexpected event should never be raised"
End Sub

Public Function CheckTestCounters()
    Dim b As Boolean
    
    b = ( _
    TestOfResultsCount > 0 And _
    TestCount > 0 And _
    TestOfCount > 0)
    
'    Debug.Assert TestOfResultsCount > 0
'    Debug.Assert TestCount > 0
'    Debug.Assert TestOfCount > 0

    CheckTestCounters = b
End Function
