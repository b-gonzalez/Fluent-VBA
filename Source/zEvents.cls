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

Private WithEvents pTestFluent As cTests
Attribute pTestFluent.VB_VarHelpID = -1
Private WithEvents pTestFluentFunction As cTests
Attribute pTestFluentFunction.VB_VarHelpID = -1

Private pFluentEvent As IFluent
Private pFluentEventOf As IFluentOf
Private pFluentEventOfResult As IFluentOf
Private pFluentEventDuplicate As IFluent

Private pFluentTestEvent As IFluent
Private pFluentFunctionTestEvent As IFluentFunction

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
    Set pTest = pFluentEvent.Meta.tests
End Property

Public Property Set setFluentOfTest(ByVal value As IFluentOf)
    Set pFluentEventOf = value
    Set pTestOf = pFluentEventOf.Meta.tests
End Property

Public Property Set setFluentTest(ByVal value As IFluent)
    Set pFluentTestEvent = value
    Set pTestFluent = pFluentTestEvent.Meta.tests
End Property

Public Property Set setFluentFunctionTest(ByVal value As IFluentFunction)
    Set pFluentFunctionTestEvent = value
    Set pTestFluentFunction = pFluentFunctionTestEvent.Meta.tests
End Property

Public Property Set setFluentEventOfResult(ByVal value As IFluentOf)
    Set pFluentEventOfResult = value
    Set pTestOfResults = pFluentEventOfResult.Meta.tests
End Property

Public Property Set setFluentEventDuplicate(ByVal value As IFluent)
    Set pFluentEventDuplicate = value
    Set pTestDuplicate = pFluentEventDuplicate.Meta.tests
End Property

Private Sub pTestFluent_TestCompleted(ByVal test As ITest)
    TestOfCount = TestOfCount + 1
    Debug.Assert pFluentTestEvent.Meta.tests.Count = TestOfCount
End Sub

Private Sub pTestFluent_TestFailed(ByVal test As ITest)
    Debug.Assert pFluentTestEvent.Meta.tests.result = test.result
    Debug.Assert pFluentTestEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestFluent_TestPassed(ByVal test As ITest)
    Debug.Assert pFluentTestEvent.Meta.tests.result = test.result
    Debug.Assert pFluentTestEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestFluentFunction_TestCompleted(ByVal test As ITest)
    TestOfCount = TestOfCount + 1
    Debug.Assert pFluentFunctionTestEvent.Meta.tests.Count = TestOfCount
End Sub

Private Sub pTestFluentFunction_TestFailed(ByVal test As ITest)
    Debug.Assert pFluentFunctionTestEvent.Meta.tests.result = test.result
    Debug.Assert pFluentFunctionTestEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestFluentFunction_TestPassed(ByVal test As ITest)
    Debug.Assert pFluentFunctionTestEvent.Meta.tests.result = test.result
    Debug.Assert pFluentFunctionTestEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestPassed(ByVal test As ITest)
    Debug.Assert pFluentEventOf.Meta.tests.result = test.result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestFailed(ByVal test As ITest)
    Debug.Assert pFluentEventOf.Meta.tests.result = test.result
    Debug.Assert pFluentEventOf.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTestOf_TestCompleted(ByVal test As ITest)
    TestOfCount = TestOfCount + 1
    Debug.Assert pFluentEventOf.Meta.tests.Count = TestOfCount
End Sub

Private Sub pTestOf_TestUnexpected(ByVal test As ITest)
'    Debug.Assert test.result = "Null" And test.HasNull
End Sub

Private Sub pTest_TestPassed(ByVal test As ITest)
    Debug.Assert pFluentEvent.Meta.tests.result = test.result
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestFailed(ByVal test As ITest)
    Debug.Assert pFluentEvent.Meta.tests.result = test.result
    Debug.Assert pFluentEvent.Meta.Printing.FluentPath = test.FluentPath
End Sub

Private Sub pTest_TestCompleted(ByVal test As ITest)
    testCount = testCount + 1
    Debug.Assert pFluentEvent.Meta.tests.Count = testCount
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
