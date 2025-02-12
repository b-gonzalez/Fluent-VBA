VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zEvents"
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
Private WithEvents pTestDuplicate As cTests
Attribute pTestDuplicate.VB_VarHelpID = -1

Private pFluentEvent As IFluent
Private pFluentEventOf As IFluentOf
Private pFluentEventOfResult As IFluentOf
Private pFluentEventDuplicate As IFluent

Private testCount As Long
Private TestOfCount As Long
Private TestOfResultsCount As Long
Private TestOfDuplicateCount As Long
Private TestDuplicateCount As Long
Private i As Long

Private Sub Class_Initialize()
    testCount = 0
    TestOfCount = 0
    TestOfResultsCount = 0
    TestDuplicateCount = 0
    TestOfDuplicateCount = 0
End Sub

Public Property Set setFluent(ByVal value As IFluent)
    Set pFluentEvent = value
    Set pTest = pFluentEvent.Meta.Tests
End Property

Public Property Set setFluentOf(ByVal value As IFluentOf)
    Set pFluentEventOf = value
    Set pTestOf = pFluentEventOf.Meta.Tests
End Property

Public Property Set setFluentEventOfResult(ByVal value As IFluentOf)
    Set pFluentEventOfResult = value
    Set pTestOfResults = pFluentEventOfResult.Meta.Tests
End Property

Public Property Set setFluentEventDuplicate(ByVal value As IFluent)
    Set pFluentEventDuplicate = value
    Set pTestDuplicate = pFluentEventDuplicate.Meta.Tests
End Property

Private Sub pTestOf_TestPassed(ByVal test As ITest)
    Debug.Assert pFluentEventOf.Meta.Tests.result = test.result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestFailed(ByVal test As ITest)
    Debug.Assert pFluentEventOf.Meta.Tests.result = test.result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestCompleted(ByVal test As ITest)
    TestOfCount = TestOfCount + 1
    Debug.Assert pFluentEventOf.Meta.Tests.Count = TestOfCount
End Sub

Private Sub pTestOf_TestUnexpected(ByVal test As ITest)
'    Debug.Assert test.result = "Null" And test.HasNull
End Sub

Private Sub pTest_TestPassed(ByVal test As ITest)
    Debug.Assert pFluentEvent.Meta.Tests.result = test.result
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestFailed(ByVal test As ITest)
    Debug.Assert pFluentEvent.Meta.Tests.result = test.result
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestCompleted(ByVal test As ITest)
    testCount = testCount + 1
    Debug.Assert pFluentEvent.Meta.Tests.Count = testCount
End Sub

Private Sub pTestOfResult_TestPassed(ByVal test As ITest)
End Sub

Private Sub pTestOfResult_TestFailed(ByVal test As ITest)
End Sub

Private Sub pTestOfResults_TestCompleted(ByVal test As ITest)
    TestOfResultsCount = TestOfResultsCount + 1
End Sub

Private Sub pTestDuplicate_DuplicateTest(ByVal test As ITest)
    TestDuplicateCount = TestDuplicateCount + 1
End Sub

Public Function CheckTestCounters() As Boolean
    Dim b As Boolean
    
    b = ( _
    TestOfResultsCount > 0 And _
    testCount > 0 And _
    TestOfCount > 0 And _
    TestDuplicateCount > 0)

    CheckTestCounters = b
End Function
