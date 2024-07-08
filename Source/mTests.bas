Attribute VB_Name = "mTests"
Option Explicit

Private mCounter As Long
Private mTestCounter As Long
Private mMiscPosTests As Long
Private mMiscNegTests As Long
Private posTestCount As Long
Private negTestCount As Long

Private mEvents As zEvents

Public Enum hw
    helloWorld
    goodbyeWorld
End Enum

Public Sub runMainTests()
    Dim fluent As IFluent
    Dim testFluent As IFluentOf
    Dim testFluentResult As IFluentOf
    Dim events As zEvents
    Dim nulTestFluent As IFluentOf
    Dim emptyTestFluent As IFluentOf
    Dim tempCounter As Long
    
    Set fluent = MakeFluent
    Set testFluentResult = MakeFluentOf
    
    Set testFluent = getAndInitTestFluent
    
    Set mEvents = getAndInitEvent(fluent, testFluent, testFluentResult)
    
    Call runEqualPosNegTests(fluent, testFluent, testFluentResult)
    
'    testFluent.Meta.Printing.PrintToSheet

    tempCounter = mCounter

    Set fluent = MakeFluent
    Set testFluent = MakeFluentOf
    
    mCounter = 0
    
    Set nulTestFluent = runNullTests(fluent, testFluent, testFluentResult)
    
    mCounter = tempCounter + nulTestFluent.Meta.Tests.Count
    
'    nulTestFluent.Meta.Printing.PrintToSheet

    tempCounter = mCounter

    Set fluent = MakeFluent
    Set testFluent = MakeFluentOf

    mCounter = 0

    Set emptyTestFluent = runEmptyTests(fluent, testFluent, testFluentResult)

    mCounter = tempCounter + emptyTestFluent.Meta.Tests.Count
    
    Set fluent = MakeFluent

    mCounter = mCounter + MiscTests(fluent)
    
'    emptyTestFluent.Meta.Printing.PrintToSheet
    
    Debug.Print "All tests Finished"
    
    Call printTestCount(mCounter + mMiscNegTests + mMiscPosTests)
    
    Call resetAndCheckCounters(mEvents, fluent, testFluent)
    
End Sub

Private Function getAndInitTestFluent() As IFluentOf
    Dim testFluent As IFluentOf
    
    Set testFluent = MakeFluentOf
    
    With testFluent.Meta
        .Printing.PassedMessage = "Success"
        .Printing.FailedMessage = "Failure"
        .Printing.UnexpectedMessage = "What?"
        
        .Tests.ToStrDev = True
    End With
    
    
    Set getAndInitTestFluent = testFluent
End Function

Private Function getAndInitEvent(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As zEvents
    
    Set mEvents = New zEvents
    
    Set mEvents.setFluent = fluent
    Set mEvents.setFluentOf = testFluent
    Set mEvents.setFluentEventOfResult = testFluentResult
    
    Set getAndInitEvent = mEvents
End Function

Private Sub runEqualPosNegTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    Dim posTestFluent As IFluentOf
    Dim negTestFluent As IFluentOf
    Dim equalTestFluent As IFluentOf
    Dim equalDict As Scripting.Dictionary
    Dim posDict As Scripting.Dictionary
    Dim negDict As Scripting.Dictionary
    Dim elem As Variant
    Dim i As Long
    Dim equalTestingInfo As ITestingFunctionsInfo
    Dim positiveTestingInfo As ITestingFunctionsInfo
    Dim negativeTestingInfo As ITestingFunctionsInfo
    Dim equalTestingInfoDict As Scripting.Dictionary
    Dim positiveTestingInfoDict As Scripting.Dictionary
    Dim negativeTestingInfoDict As Scripting.Dictionary
    
    fluent.Meta.Printing.Category = "Fluent - EqualityTests"
    testFluent.Meta.Printing.Name = "Test Fluent - abc 123"
    testFluent.Meta.Printing.Category = "Test Fluent - EqualityTests"
    Set equalTestFluent = EqualityDocumentationTests(fluent, testFluent, testFluentResult)
    Set equalTestingInfo = equalTestFluent.Meta.Tests.testingInfo
    Set equalTestingInfoDict = equalTestingInfo.TestFuncInfoToDict
    
'    Call equalTestingInfo.validateTfiDictCounters(equalTestingInfoDict)

'    Call equalTestingInfo.printTestFunInfo
'
'    Stop
    
    'The equality tests cannot use validateTestDict counter since it will fail since they
    'only run the equality tests
    
    testFluent.Meta.Tests.resetTestingInfo

    fluent.Meta.Printing.Category = "Fluent - positiveDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - positiveDocumentationTests"
    Set posTestFluent = positiveDocumentationTests(fluent, testFluent, testFluentResult)
    
    Set positiveTestingInfo = posTestFluent.Meta.Tests.testingInfo
    Set positiveTestingInfoDict = positiveTestingInfo.TestFuncInfoToDict
    
    Debug.Assert positiveTestingInfo.validateTfiDictCounters(positiveTestingInfoDict)

'    Call positiveTestingInfo.printTestFunInfo
'    Stop
    
    testFluent.Meta.Tests.resetTestingInfo
    
    fluent.Meta.Printing.Category = "Fluent - negativeDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - negativeDocumentationTests"
    Set negTestFluent = negativeDocumentationTests(fluent, testFluent, testFluentResult)
    
    Set negativeTestingInfo = negTestFluent.Meta.Tests.testingInfo
    Set negativeTestingInfoDict = negativeTestingInfo.TestFuncInfoToDict
    
    Debug.Assert negativeTestingInfo.validateTfiDictCounters(negativeTestingInfoDict)
    
'    Stop
    
    Debug.Assert testFluent.Meta.Tests.testingInfo.TestFuncInfoDictsAreIdentical(positiveTestingInfoDict, negativeTestingInfoDict)
    
'    Call negativeTestingInfo.printTestFunInfo
'
'    Stop
    
'    negFluentOfStr = getFluentOfCounts(negTestFluent)

    Debug.Assert posTestCount = negTestCount
    
    With posTestFluent.Meta
        For i = 1 To .Tests.Count
            Debug.Assert .Tests(i).functionName = negTestFluent.Meta.Tests(i).functionName
        Next i
    End With
    
    Debug.Assert validateNegativeCounters(testFluent)
End Sub

Private Function runNullTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    Dim nulTestFluent As IFluentOf
    Dim tempDict As Scripting.Dictionary
    Dim posNullTestInfo As ITestingFunctionsInfo
    Dim negNullTestInfo As ITestingFunctionsInfo
    Dim posNullDict As Scripting.Dictionary
    Dim negNullDict As Scripting.Dictionary
    
    Set tempDict = New Scripting.Dictionary
    
    tempDict("count") = 1
    
    fluent.Meta.Printing.Category = "Fluent - nullDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - positiveNullDocumentationTests"
    Set nulTestFluent = positiveNullDocumentationTests(fluent, testFluent, testFluentResult)
    Set posNullTestInfo = nulTestFluent.Meta.Tests.testingInfo
    Set posNullDict = posNullTestInfo.TestFuncInfoToDict
    
    Debug.Assert posNullTestInfo.validateTfiDictCounters(posNullDict, 2) 'set to 2 for OneOf and SameTypeAs which do not have null tests
    
'    Stop
    
    nulTestFluent.Meta.Tests.resetTestingInfo
    
    fluent.Meta.Printing.Category = "Fluent - nullDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - negativeNullDocumentationTests"
    Set nulTestFluent = negativeNullDocumentationTests(fluent, testFluent, testFluentResult)
    Set negNullTestInfo = nulTestFluent.Meta.Tests.testingInfo
    Set negNullDict = negNullTestInfo.TestFuncInfoToDict
    
    Debug.Assert negNullTestInfo.validateTfiDictCounters(negNullDict, 2) 'set to 2 for OneOf and SameTypeAs which do not have null tests
    
'    Stop
    
    Debug.Assert nulTestFluent.Meta.Tests.testingInfo.TestFuncInfoDictsAreIdentical(posNullDict, negNullDict)
    
    Set runNullTests = nulTestFluent
End Function

Private Function runEmptyTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    Dim emptyTestFluent As IFluentOf
    Dim tempDict As Scripting.Dictionary
    Dim emptyDict As Scripting.Dictionary
    
    Set tempDict = New Scripting.Dictionary
    
    fluent.Meta.Printing.Category = "Fluent - emptyDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - emptyDocumentationTests"
    Set emptyTestFluent = emptyDocumentationTests(fluent, testFluent, testFluentResult)
    
    Set runEmptyTests = emptyTestFluent
End Function

Private Sub resetAndCheckCounters(events As zEvents, fluent As IFluent, testFluent As IFluentOf)
    mCounter = 0
    
    mTestCounter = 0
    
    mMiscNegTests = 0
    
    mMiscPosTests = 0

    Debug.Assert events.CheckTestCounters

    Debug.Assert checkResetCounters(fluent, testFluent)
End Sub

Private Sub printTestCount(testCount As Long)
    If testCount > 1 Then
        Debug.Print testCount & " tests finished!" & vbNewLine
    ElseIf testCount = 1 Then
        Debug.Print "1 Test finished!"
    End If
End Sub

Private Sub TrueAssertAndRaiseEvents(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    Dim td As ITestDev
    Dim inputIter As String
    Dim inputRecur As String
    Dim valueIter As String
    Dim valueRecur As String

    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.Tests.Count = mCounter

    With fluent.Meta.Tests
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    If testFluent.Meta.Tests.ToStrDev Then
        With testFluent.Meta.Tests
            Set td = .Item(.Count)
        End With
    
        inputIter = td.TestInputIter
        inputRecur = td.TestInputRecur
        valueIter = td.TestValueIter
        valueRecur = td.TestValueRecur
        
        Debug.Assert _
        inputIter = inputRecur And _
        valueIter = valueRecur
    End If
End Sub

Private Sub FalseAssertAndRaiseEvents(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    Dim td As ITestDev
    Dim inputIter As String
    Dim inputRecur As String
    Dim valueIter As String
    Dim valueRecur As String
    
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.Tests.Count = mCounter

    With fluent.Meta.Tests
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    If testFluent.Meta.Tests.ToStrDev Then
        With testFluent.Meta.Tests
            Set td = .Item(.Count)
        End With
    
        inputIter = td.TestInputIter
        inputRecur = td.TestInputRecur
        valueIter = td.TestValueIter
        valueRecur = td.TestValueRecur
        
        Debug.Assert _
        inputIter = inputRecur And _
        valueIter = valueRecur
    End If
End Sub

Private Sub NullAssertAndRaiseEvents(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.Tests.Count = mCounter

'    If IsNull(fluent.TestValue) Then
        With fluent
            Debug.Assert testFluentResult.Of(.TestValue).Should.Be.EqualTo(Null)
            Debug.Assert testFluentResult.Of(.TestValue).ShouldNot.Be.EqualTo(True)
            Debug.Assert testFluentResult.Of(.TestValue).ShouldNot.Be.EqualTo(False)
        End With
'    End If
End Sub

Private Sub EmptyAssertAndRaiseEvents(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.Tests.Count = mCounter

'    If IsNull(fluent.TestValue) Then
        With fluent
            Debug.Assert Information.IsEmpty(fluent.TestValue)
'            Debug.Assert fluent.TestValue = ""
'            Debug.Assert fluent.TestValue = 0
        End With
'    End If
End Sub

Private Function EqualityDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    Dim test As ITest
    Dim i As Long
    Dim resultBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim counter As Long
    
    counter = 0

    With fluent.Meta.Tests
    
        fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(False).Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(False).ShouldNot.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).ShouldNot.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).ShouldNot.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).ShouldNot.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(0).ShouldNot.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Approximate equality tests
    
        testFluent.Meta.Tests.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("TRUE").ShouldNot.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("TRUE").ShouldNot.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("FALSE").ShouldNot.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("FALSE").ShouldNot.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("true").ShouldNot.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("true").ShouldNot.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("false").ShouldNot.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("false").ShouldNot.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.ApproximateEqual = False
        
        '//Null and Empty tests
        
        fluent.TestValue = testFluent.Of(Null).Should.Be.EqualTo(Null)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.EqualTo(Null)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
'        fluent.TestValue = testFluent.Of().Should.Be.EqualTo(Empty)
'        Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
'
'        fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo(Empty)
'        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Empty).Should.Be.EqualTo(Empty)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Empty).ShouldNot.Be.EqualTo(Empty)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("").Should.Be.EqualTo(Empty)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("").ShouldNot.Be.EqualTo(Empty)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(Empty)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).ShouldNot.Be.EqualTo(Empty)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.Be.EqualTo(Empty)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.Be.EqualTo(Empty)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    End With
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.Result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).Result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            resultBool = test.Result = .Tests(i).Result
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.testingValue = .Tests(i).testingValue
            inputBool = test.testingInput = .Tests(i).testingInput
            
            Debug.Assert resultBool And fluentBool And valueBool And inputBool
            
            i = i + 1
        Next test
    End With
    
    Debug.Print "Equality tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set EqualityDocumentationTests = testFluent
End Function

Private Function positiveDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    'Dim testFluent As cFluentOf
    Dim test As ITest
    Dim col As Collection
    Dim col2 As Collection
    Dim col3 As Collection
    Dim arr As Variant
    Dim arr2 As Variant
    Dim d As Object
    Dim al As Object
    Dim i As Long
    Dim counter As Long
    Dim resultBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim tfRecur As IFluentOf
    Dim tfIter As IFluentOf
    Dim b As Boolean
    
    With fluent.Meta.Tests
        fluent.TestValue = testFluent.Of("""abc""").Should.Be.EqualTo("""abc""")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).Should.Be.EqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(0)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of("10").Should.Contain("1")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.Contain("0")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.Contain("10")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.Contain("2")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello world").Should.Contain("Hello")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello world").Should.Contain("world")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Contain("ru")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").Should.Contain("1")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").Should.Contain("0")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").Should.Contain("10")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").Should.Contain("2")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""Hello world""").Should.Contain("Hello")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""Hello world""").Should.Contain("world")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Contain("ru")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.Contain("als")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.StartWith(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).Should.StartWith(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.StartWith("1")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("10").Should.StartWith("2")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello World").Should.StartWith("Hello")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        
        fluent.TestValue = testFluent.Of("1 ""0"" ").Should.StartWith("1")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("1 ""0"" ").Should.StartWith("2")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello ""World"" ").Should.StartWith("Hello")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        
        
        fluent.TestValue = testFluent.Of(True).Should.StartWith("True")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.StartWith("T")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.StartWith("False")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.StartWith("F")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.EndWith(0)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.EndWith(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.EndWith("0")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.EndWith("2")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello World").Should.EndWith("World")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of(" ""1"" 0").Should.EndWith("0")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""1"" 0").Should.EndWith("2")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""Hello"" World").Should.EndWith("World")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.EndWith("True")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.EndWith("e")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.EndWith("False")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.EndWith("e")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").Should.Have.LengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Have.LengthOf(4)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(Len("10"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").Should.Have.LengthOf(Len("abc"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Have.LengthOf(Len("True"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of("10").Should.Have.MaxLengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.Have.MaxLengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Have.MaxLengthOf(Len("True"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.Have.MaxLengthOf(Len("False"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MinLengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MinLengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.Have.MinLengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").Should.Have.MinLengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.Have.MinLengthOf(Len("True"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.Have.MinLengthOf(Len("False"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).Should.Be.GreaterThanOrEqualTo(9.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).Should.Be.GreaterThanOrEqualTo(11.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).Should.Be.LessThanOrEqualTo(10.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).Should.Be.LessThanOrEqualTo(11.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).Should.Be.LessThanOrEqualTo(9.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(10, 10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(9.99, 10.01)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(9, 11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(9.1, 11.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(11, 9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(11.1, 9.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(1, 3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(0, 2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(2, 2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(3, 1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(2, 0)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(9, 10, 11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(9, 11, 13)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf() 'intentionally empty
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ' //Object and data structure tests
        
        Set col = New Collection
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(col).Should.Be.OneOf(col, d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(col, d, 10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(col).Should.Be.Something
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = Nothing
        fluent.TestValue = testFluent.Of(col).Should.Be.Something
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        

'InDataStructure

        Set tfRecur = MakeFluentOf
        Set tfIter = MakeFluentOf
        
        tfRecur.Meta.Tests.Algorithm = flRecursive
        tfIter.Meta.Tests.Algorithm = flIterative
        
        arr = Array()
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(False) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1, 1)
        arr(0, 0, 0) = 6
        arr(0, 0, 1) = 7
        arr(0, 1, 0) = 8
        arr(0, 1, 1) = 9
        arr(1, 0, 0) = 10
        arr(1, 0, 1) = 11
        arr(1, 1, 0) = 12
        arr(1, 1, 1) = 13
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(col) = tfIter.Of(10).Should.Be.InDataStructure(col)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(col) = tfIter.Of(10).Should.Be.InDataStructure(col)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(d) = tfIter.Of(10).Should.Be.InDataStructure(d)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(d) = tfIter.Of(10).Should.Be.InDataStructure(d)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.Keys) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(d.Keys) = tfIter.Of(10).Should.Be.InDataStructure(d.Keys)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructure(al) = tfIter.Of(10).Should.Be.InDataStructure(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

'InDataStructures
        
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(arr2) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(arr2) = tfIter.Of(10).Should.Be.InDataStructures(arr2)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            
        ReDim arr(1, 1)
        arr(0, 0) = 12
        arr(0, 1) = 13
        arr(1, 0) = 14
        arr(1, 1) = 15
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(12).Should.Be.InDataStructures(arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(12).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(12).Should.Be.InDataStructures(arr, arr2)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        ReDim arr(1, 1, 1)
        arr(0, 0, 0) = 6
        arr(0, 0, 1) = 7
        arr(0, 1, 0) = 8
        arr(0, 1, 1) = 9
        arr(1, 0, 0) = 10
        arr(1, 0, 1) = 11
        arr(1, 1, 0) = 12
        arr(1, 1, 1) = 13
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(9).Should.Be.InDataStructures(arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(9).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(9).Should.Be.InDataStructures(arr, arr2)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(10).Should.Be.InDataStructures(arr, arr2)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        fluent.TestValue = testFluent.Of(13).Should.Be.InDataStructures(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(13).Should.Be.InDataStructures(col) = tfIter.Of(13).Should.Be.InDataStructures(col)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(16).Should.Be.InDataStructures(arr, col, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(16).Should.Be.InDataStructures(arr, col, arr2) = tfIter.Of(16).Should.Be.InDataStructures(arr, col, arr2)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(col) = tfIter.Of(10).Should.Be.InDataStructures(col)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        arr = Array(12, 13, 14)
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(14).Should.Be.InDataStructures(col, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(14).Should.Be.InDataStructures(col, arr) = tfIter.Of(14).Should.Be.InDataStructures(col, arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
    
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(d) = tfIter.Of(10).Should.Be.InDataStructures(d)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(2).Should.Be.InDataStructures(d.Items, d.Keys)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(2).Should.Be.InDataStructures(d.Items, d.Keys) = tfIter.Of(2).Should.Be.InDataStructures(d.Items, d.Keys)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = Nothing
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(d) = tfIter.Of(10).Should.Be.InDataStructures(d)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(6, Array(7, Array(8)))
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(8).Should.Be.InDataStructures(al, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(8).Should.Be.InDataStructures(al, arr) = tfIter.Of(8).Should.Be.InDataStructures(al, arr)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        ' //Approximate equality tests
        testFluent.Meta.Tests.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("10").Should.Be.EqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("True").Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//default epsilon for double comparisons is 0.000001
        '//the default can be modified by setting a value
        '//for the epsilon property in the Meta object.
        
        fluent.TestValue = testFluent.Of(5.0000001).Should.Be.EqualTo(5)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Evaluation tests
        
        fluent.TestValue = testFluent.Of(True).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).Should.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).Should.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("true").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("false").Should.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("TRUE").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("FALSE").Should.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).Should.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).Should.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).Should.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("-1").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("-1").Should.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("0").Should.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("0").Should.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(5 + 5).Should.EvaluateTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("5 + 5").Should.EvaluateTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("5 + 5 = 10").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("5 + 5 > 9").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array()
        fluent.TestValue = testFluent.Of(Information.TypeName(arr) = "Variant()").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Information.IsArray(arr)).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(Information.TypeName(col) = "Collection").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf col Is Collection).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(Information.TypeName(d) = "Dictionary").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf d Is Scripting.Dictionary).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Testing errors is possible if they're put in strings
        fluent.TestValue = testFluent.Of("1 / 0").Should.EvaluateTo(CVErr(xlErrDiv0))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'//Alphabetic
        
        fluent.TestValue = testFluent.Of("abc").Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc!@#").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("!@#").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = True
        
        fluent.TestValue = testFluent.Of("abc def").Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" abc ").Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" abc!@# ").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" 123 ").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" !@# ").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc""").Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc!@#""").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123""").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""!@#""").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc"" ").Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc!@#"" ").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123"" ").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""!@#"" ").Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc """).Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc!@# """).Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123 """).Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" !@# """).Should.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = False
        
        fluent.TestValue = testFluent.Of(123).Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123").Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123!@#").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("!@#").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = True
        
        fluent.TestValue = testFluent.Of("123 456").Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123""").Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123!@#""").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc""").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""!@#""").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123"" ").Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123!@#"" ").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc"" ").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""!@#"" ").Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123 """).Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123!@# """).Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc """).Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" !@# """).Should.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = False
        
        fluent.TestValue = testFluent.Of("abc123").Should.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("!@#").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = True
        
        fluent.TestValue = testFluent.Of("abc 123").Should.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc123""").Should.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc""").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123""").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""!@#""").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc123"" ").Should.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc"" ").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123"" ").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""!@#"" ").Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc123 """).Should.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc """).Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123 """).Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" !@# """).Should.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = False
        
        fluent.TestValue = testFluent.Of(CStr(Excel.Evaluate("1 / 0"))).Should.Be.EqualTo("Error 2007")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("1 / 0").Should.Be.Erroneous
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        On Error Resume Next
            Debug.Print 1 / 0
            
            fluent.TestValue = testFluent.Of(Err).Should.Be.Erroneous
        On Error GoTo 0
        
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("1 / 0").Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        On Error Resume Next
            Debug.Print 1 / 0
            
            fluent.TestValue = testFluent.Of(Err).Should.Have.ErrorDescriptionOf("Division by zero")
        On Error GoTo 0
        
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("1 / 0").Should.Have.ErrorNumberOf(2007)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        On Error Resume Next
            Debug.Print 1 / 0
            
            fluent.TestValue = testFluent.Of(Err).Should.Have.ErrorNumberOf(11)
        On Error GoTo 0
        
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CBool(True)).Should.Have.SameTypeAs(CBool(True))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("Hello World!")).Should.Have.SameTypeAs(CStr("Goodbye World!"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("""Hello World!""")).Should.Have.SameTypeAs(CStr("""Goodbye World!"""))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of(CStr("""Hello World!""")).Should.Have.SameTypeAs(CStr("Goodbye World!"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("Hello World!")).Should.Have.SameTypeAs(CStr("""Goodbye World!"""))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CLng(12345)).Should.Have.SameTypeAs(CLng(54321))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CSng(123.45)).Should.Have.SameTypeAs(CSng(543.21))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CDbl(123.45)).Should.Have.SameTypeAs(CDbl(543.21))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CDate(#12/31/1999#)).Should.Have.SameTypeAs(CDate(#12/31/2000#))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'        fluent.TestValue = testFluent.Of(CLng(hw.helloWorld)).Should.Have.SameTypeAs(CLng(hw.goodbyeWorld))
'        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameTypeAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Nothing).Should.Have.SameTypeAs(Nothing)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(col).Should.Have.SameTypeAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        fluent.TestValue = testFluent.Of(col).Should.Have.SameTypeAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(d).Should.Have.SameTypeAs(d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        fluent.TestValue = testFluent.Of(d).Should.Have.SameTypeAs(d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CLng(123)).Should.Have.SameTypeAs(CStr("Hello world"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CLng(123)).Should.Have.SameTypeAs(CDbl(123.456))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(CLng(123)).Should.Have.SameTypeAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1
        
        With testFluent.Of(col).Should.Be
            fluent.TestValue = .IdenticalTo(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 2
        col2.Add 1
        col3.Add 1
        
        With testFluent.Of(col).Should.Be
            fluent.TestValue = .IdenticalTo(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 2
        col3.Add 1
        
        fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = testFluent.Of(col2).Should.Be.IdenticalTo(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = testFluent.Of(col2).Should.Be.IdenticalTo(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col2).Should.Be.IdenticalTo(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1
        With testFluent.Of(col2).Should.Be
            fluent.TestValue = .IdenticalTo(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        Set col = Nothing
        Set col2 = Nothing
        Set col3 = Nothing
        
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).Should.Be.IdenticalTo(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        fluent.TestValue = testFluent.Of(arr).Should.Be.IdenticalTo(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        
        With testFluent.Of(arr).Should.Be
            fluent.TestValue = .IdenticalTo(arr2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(Array(1, 2, 3))
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 3, 4)
        fluent.TestValue = testFluent.Of(arr).Should.Be.IdenticalTo(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        
        With testFluent.Of(Array(2, 3, 4)).Should.Be
            fluent.TestValue = .IdenticalTo(arr)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(arr2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1
        
        With testFluent.Of(col).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 2
        col2.Add 1
        col3.Add 1

        With testFluent.Of(col).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 2
        col3.Add 1

        With testFluent.Of(col2).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 2

        With testFluent.Of(col3).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col2).Should.Have.ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1

        With testFluent.Of(col2).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        Set col = Nothing
        Set col2 = Nothing
        Set col3 = Nothing
        
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).Should.Have.ExactSameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        fluent.TestValue = testFluent.Of(arr).Should.Have.ExactSameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr

        With testFluent.Of(arr).Should.Have
            fluent.TestValue = .ExactSameElementsAs(arr2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(Array(1, 2, 3))
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 3, 4)
        fluent.TestValue = testFluent.Of(arr).Should.Have.ExactSameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr

        With testFluent.Of(Array(2, 3, 4)).Should.Have
            fluent.TestValue = .ExactSameElementsAs(arr)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(arr2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(Array(1))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(Array(2))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1

        With testFluent.Of(Array(1)).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col3 = New Collection
        col.Add 2
        col2.Add 1

        With testFluent.Of(col).Should.Have
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2

        With testFluent.Of(col2).Should.Have
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 2

        With testFluent.Of(col2).Should.Have
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col2).Should.Have.ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        

        Set col = New Collection
        col.Add 1
        With testFluent.Of(Array(2)).Should.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        Set col = Nothing
        
        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        col2.Add 2
        col2.Add 1
        col.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1)
        arr2 = Array(1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 1)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 1)
        arr2 = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        arr = Array(1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        arr = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 1
        col2.Add 3
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 2
        col2.Add 1
        col.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1)
        arr2 = Array(2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 1, 0)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        arr = Array(2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        


        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(1, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 2
        col.Add 2
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1, 0)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        



        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 1
        col2.Add 2
        col2.Add 3
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


        arr = Array(1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        arr = Array(1)
        arr2 = Array(1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


        arr = Array(1, 2)
        arr2 = Array(1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        arr = Array(1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 1
        col2.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 3
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        arr = Array(1)
        arr2 = Array(2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


        arr = Array(1, 2)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(3, 2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        arr = Array(2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(3, 2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    End With
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.Result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).Result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            resultBool = test.Result = .Tests(i).Result
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput
            
            Debug.Assert resultBool And fluentBool And valueBool And inputBool
            
            i = i + 1
        Next test
    End With
    
    Debug.Print "Positive tests finished"
    posTestCount = mTestCounter + mMiscPosTests
    mTestCounter = 0
    printTestCount (posTestCount)
    
    Set positiveDocumentationTests = testFluent
    
End Function

Private Function negativeDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    'Dim testFluent As cFluentOf
    Dim test As ITest
    Dim col As Collection
    Dim col2 As Collection
    Dim col3 As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    Dim i As Long
    Dim arr2 As Variant
    Dim resultBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim tfRecur As IFluentOf
    Dim tfIter As IFluentOf
    Dim b As Boolean
    
    
    
    
    With fluent.Meta.Tests
    
        fluent.TestValue = testFluent.Of("""abc""").ShouldNot.Be.EqualTo("""abc""")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.EqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(0)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Contain("1")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Contain("0")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Contain("10")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Contain("2")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Contain("Hello")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Contain("world")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").ShouldNot.Contain("1")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").ShouldNot.Contain("0")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").ShouldNot.Contain("10")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""10""").ShouldNot.Contain("2")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""Hello world""").ShouldNot.Contain("Hello")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""Hello world""").ShouldNot.Contain("world")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Contain("ru")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.Contain("als")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abcde""").ShouldNot.Contain("abc")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.StartWith("1")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.StartWith("2")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of("Hello World").ShouldNot.StartWith("Hello")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("1 ""0"" ").ShouldNot.StartWith("1")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of("1 ""0"" ").ShouldNot.StartWith("2")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello ""World"" ").ShouldNot.StartWith("Hello")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.StartWith("True")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.StartWith("T")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.StartWith("False")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.StartWith("F")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(0)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.EndWith("0")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.EndWith("2")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("Hello World").ShouldNot.EndWith("World")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of(" ""1"" 0").ShouldNot.EndWith("0")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""1"" 0").ShouldNot.EndWith("2")
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""Hello"" World").ShouldNot.EndWith("World")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.EndWith("True")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.EndWith("e")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.EndWith("False")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.EndWith("e")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Have.LengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Have.LengthOf(4)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(Len("10"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Have.LengthOf(Len("abc"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Have.LengthOf(Len("True"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Have.MaxLengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Have.MaxLengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Have.MaxLengthOf(Len("True"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.Have.MaxLengthOf(Len("False"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MinLengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MinLengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Have.MinLengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("10").ShouldNot.Have.MinLengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.Have.MinLengthOf(Len("True"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.Have.MinLengthOf(Len("False"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).ShouldNot.Be.GreaterThanOrEqualTo(9.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).ShouldNot.Be.GreaterThanOrEqualTo(11.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).ShouldNot.Be.LessThanOrEqualTo(10.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).ShouldNot.Be.LessThanOrEqualTo(11.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10.1).ShouldNot.Be.LessThanOrEqualTo(9.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(10, 10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(9.99, 10.01)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(9, 11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(9.1, 11.1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(11, 9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(11.1, 9.1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(1, 3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(0, 2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(2, 2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(3, 1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(2, 0)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(9, 10, 11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(9, 11, 13)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf()  'intentionally empty
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ' //Object and data structure tests
        
        Set col = New Collection
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.OneOf(col, d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(col, d, 10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Something
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = Nothing
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Something
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'InDataStructure

        Set tfRecur = MakeFluentOf
        Set tfIter = MakeFluentOf
        
        tfRecur.Meta.Tests.Algorithm = flRecursive
        tfIter.Meta.Tests.Algorithm = flIterative
        
        arr = Array()
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(False) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        ReDim arr(1, 1, 1)
        arr(0, 0, 0) = 6
        arr(0, 0, 1) = 7
        arr(0, 1, 0) = 8
        arr(0, 1, 1) = 9
        arr(1, 0, 0) = 10
        arr(1, 0, 1) = 11
        arr(1, 1, 0) = 12
        arr(1, 1, 1) = 13
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(col) = tfIter.Of(10).ShouldNot.Be.InDataStructure(col)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(col) = tfIter.Of(10).ShouldNot.Be.InDataStructure(col)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(d) = tfIter.Of(10).ShouldNot.Be.InDataStructure(d)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(d) = tfIter.Of(10).ShouldNot.Be.InDataStructure(d)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d.Keys) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(d.Keys) = tfIter.Of(10).ShouldNot.Be.InDataStructure(d.Keys)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(al) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(al) = tfIter.Of(10).ShouldNot.Be.InDataStructure(al)
        mMiscNegTests = mMiscNegTests + 1
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

'InDataStructures

        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(arr2) 'with implicit recur
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(arr2) = tfIter.Of(10).ShouldNot.Be.InDataStructures(arr2)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            
        ReDim arr(1, 1)
        arr(0, 0) = 12
        arr(0, 1) = 13
        arr(1, 0) = 14
        arr(1, 1) = 15
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(12).ShouldNot.Be.InDataStructures(arr, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(12).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(12).ShouldNot.Be.InDataStructures(arr, arr2)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
        ReDim arr(1, 1, 1)
        arr(0, 0, 0) = 6
        arr(0, 0, 1) = 7
        arr(0, 1, 0) = 8
        arr(0, 1, 1) = 9
        arr(1, 0, 0) = 10
        arr(1, 0, 1) = 11
        arr(1, 1, 0) = 12
        arr(1, 1, 1) = 13
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(9).ShouldNot.Be.InDataStructures(arr, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(9).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(9).ShouldNot.Be.InDataStructures(arr, arr2)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(arr, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(10).ShouldNot.Be.InDataStructures(arr, arr2)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        fluent.TestValue = testFluent.Of(13).ShouldNot.Be.InDataStructures(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(13).ShouldNot.Be.InDataStructures(col) = tfIter.Of(13).ShouldNot.Be.InDataStructures(col)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2) = tfIter.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)



        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(col) = tfIter.Of(10).ShouldNot.Be.InDataStructures(col)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set col = Nothing

        arr = Array(12, 13, 14)
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(14).ShouldNot.Be.InDataStructures(col, arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(14).ShouldNot.Be.InDataStructures(col, arr) = tfIter.Of(14).ShouldNot.Be.InDataStructures(col, arr)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set col = Nothing

        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(d) = tfIter.Of(10).ShouldNot.Be.InDataStructures(d)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing ''

        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(2).ShouldNot.Be.InDataStructures(d.Items, d.Keys)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(2).ShouldNot.Be.InDataStructures(d.Items, d.Keys) = tfIter.Of(2).ShouldNot.Be.InDataStructures(d.Items, d.Keys)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set d = Nothing
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(d) = tfIter.Of(10).ShouldNot.Be.InDataStructures(d)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
        Set d = Nothing

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        arr = Array(6, Array(7, Array(8)))
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(8).ShouldNot.Be.InDataStructures(al, arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(8).ShouldNot.Be.InDataStructures(al, arr) = tfIter.Of(8).ShouldNot.Be.InDataStructures(al, arr)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
        mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
        fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)



    
        ' //Approximate equality tests
        testFluent.Meta.Tests.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("10").ShouldNot.Be.EqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("True").ShouldNot.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//default epsilon for double comparisons is 0.000001
        '//the default can be modified by setting a value
        '//for the epsilon property in the Meta object.
        
        fluent.TestValue = testFluent.Of(5.0000001).ShouldNot.Be.EqualTo(5)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Evaluation tests
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("true").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("false").ShouldNot.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("TRUE").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("FALSE").ShouldNot.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).ShouldNot.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).ShouldNot.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).ShouldNot.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("-1").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("-1").ShouldNot.EvaluateTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("0").ShouldNot.EvaluateTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("0").ShouldNot.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(5 + 5).ShouldNot.EvaluateTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("5 + 5").ShouldNot.EvaluateTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("5 + 5 = 10").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("5 + 5 > 9").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array()
        fluent.TestValue = testFluent.Of(Information.TypeName(arr) = "Variant()").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Information.IsArray(arr)).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(Information.TypeName(col) = "Collection").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf col Is Collection).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(Information.TypeName(d) = "Dictionary").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf d Is Scripting.Dictionary).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Testing errors is possible if they're put in strings
        fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.EvaluateTo(CVErr(xlErrDiv0))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc!@#").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("!@#").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        testFluent.Meta.Tests.CleanTestStrings = True
        
        
        fluent.TestValue = testFluent.Of("abc def").ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" abc ").ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" abc!@# ").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" 123 ").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" !@# ").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of("""abc""").ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc!@#""").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123""").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""!@#""").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc"" ").ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc!@#"" ").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123"" ").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""!@#"" ").ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc """).ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc!@# """).ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123 """).ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" !@# """).ShouldNot.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = False
        
        fluent.TestValue = testFluent.Of(123).ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123").ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123!@#").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("!@#").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = True
        
        fluent.TestValue = testFluent.Of("123 456").ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123""").ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123!@#""").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of("""abc""").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""!@#""").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123"" ").ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123!@#"" ").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of(" ""abc"" ").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""!@#"" ").ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123 """).ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123!@# """).ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                
        fluent.TestValue = testFluent.Of(""" abc """).ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" !@# """).ShouldNot.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = False
        
        fluent.TestValue = testFluent.Of("abc123").ShouldNot.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("123").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("!@#").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = True
        
        fluent.TestValue = testFluent.Of("abc 123").ShouldNot.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc123""").ShouldNot.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""abc""").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""123""").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("""!@#""").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc123"" ").ShouldNot.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""abc"" ").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""123"" ").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(" ""!@#"" ").ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc123 """).ShouldNot.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" abc """).ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" 123 """).ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(""" !@# """).ShouldNot.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.Tests.CleanTestStrings = False
        
        fluent.TestValue = testFluent.Of(CStr(Excel.Evaluate("1 / 0"))).ShouldNot.Be.EqualTo("Error 2007")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        On Error Resume Next
            Debug.Print 1 / 0
            
            fluent.TestValue = testFluent.Of(Err).ShouldNot.Be.Erroneous
        On Error GoTo 0
        
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.Be.Erroneous
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        On Error Resume Next
            Debug.Print 1 / 0
            
            fluent.TestValue = testFluent.Of(Err).ShouldNot.Have.ErrorDescriptionOf("Division by zero")
        On Error GoTo 0
        
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.Have.ErrorNumberOf(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        On Error Resume Next
            Debug.Print 1 / 0
            
            fluent.TestValue = testFluent.Of(Err).ShouldNot.Have.ErrorNumberOf(11)
        On Error GoTo 0
        
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of(CBool(True)).ShouldNot.Have.SameTypeAs(CBool(True))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("Hello World!")).ShouldNot.Have.SameTypeAs(CStr("Goodbye World!"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("""Hello World!""")).ShouldNot.Have.SameTypeAs(CStr("""Goodbye World!"""))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("""Hello World!""")).ShouldNot.Have.SameTypeAs(CStr("Goodbye World!"))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CStr("Hello World!")).ShouldNot.Have.SameTypeAs(CStr("""Goodbye World!"""))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CLng(12345)).ShouldNot.Have.SameTypeAs(CLng(54321))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CSng(123.45)).ShouldNot.Have.SameTypeAs(CSng(543.21))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CDbl(123.45)).ShouldNot.Have.SameTypeAs(CDbl(543.21))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CDate(#12/31/1999#)).ShouldNot.Have.SameTypeAs(CDate(#12/31/2000#))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'        fluent.TestValue = testFluent.Of(CLng(hw.helloWorld)).ShouldNot.Have.SameTypeAs(CLng(hw.goodbyeWorld))
'        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameTypeAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Nothing).ShouldNot.Have.SameTypeAs(Nothing)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameTypeAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameTypeAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameTypeAs(d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameTypeAs(d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        fluent.TestValue = testFluent.Of(CLng(123)).ShouldNot.Have.SameTypeAs(CStr("Hello world"))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(CLng(123)).ShouldNot.Have.SameTypeAs(CDbl(123.456))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(CLng(123)).ShouldNot.Have.SameTypeAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing

        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1
        
        With testFluent.Of(col).ShouldNot.Be
            fluent.TestValue = .IdenticalTo(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 2
        col2.Add 1
        col3.Add 1
        
        With testFluent.Of(col).ShouldNot.Be
            fluent.TestValue = .IdenticalTo(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 2
        col3.Add 1
        
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = testFluent.Of(col2).ShouldNot.Be.IdenticalTo(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = testFluent.Of(col2).ShouldNot.Be.IdenticalTo(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col2).ShouldNot.Be.IdenticalTo(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1
        With testFluent.Of(col2).ShouldNot.Be
            fluent.TestValue = .IdenticalTo(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        Set col = Nothing
        Set col2 = Nothing
        Set col3 = Nothing
        
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Be.IdenticalTo(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Be.IdenticalTo(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        
        With testFluent.Of(arr).ShouldNot.Be
            fluent.TestValue = .IdenticalTo(arr2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(Array(1, 2, 3))
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 3, 4)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Be.IdenticalTo(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        
        With testFluent.Of(Array(2, 3, 4)).ShouldNot.Be
            fluent.TestValue = .IdenticalTo(arr)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .IdenticalTo(arr2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1
        
        With testFluent.Of(col).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 2
        col2.Add 1
        col3.Add 1

        With testFluent.Of(col).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 2
        col3.Add 1

        With testFluent.Of(col2).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 2

        With testFluent.Of(col3).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col2).ShouldNot.Have.ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 1
        col3.Add 1

        With testFluent.Of(col2).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col3)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        Set col = Nothing
        Set col2 = Nothing
        Set col3 = Nothing
        
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.ExactSameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.ExactSameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr

        With testFluent.Of(arr).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(arr2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(Array(1, 2, 3))
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 3, 4)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.ExactSameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = arr

        With testFluent.Of(Array(2, 3, 4)).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(arr)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(arr2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(Array(1))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(Array(2))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1

        With testFluent.Of(Array(1)).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col3 = New Collection
        col.Add 2
        col2.Add 1

        With testFluent.Of(col).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col2)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2

        With testFluent.Of(col2).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col3 = New Collection
        col.Add 1
        col2.Add 2

        With testFluent.Of(col2).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col2).ShouldNot.Have.ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        With testFluent.Of(Array(2)).ShouldNot.Have
            fluent.TestValue = .ExactSameElementsAs(col)
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            fluent.TestValue = .ExactSameElementsAs(Array(1))
            Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        End With
        Set col = Nothing
        
        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        col2.Add 2
        col2.Add 1
        col.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1)
        arr2 = Array(1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 1)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 1)
        arr2 = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        arr = Array(1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 1
        arr = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 1
        col2.Add 3
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 2
        col2.Add 1
        col.Add 2
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1)
        arr2 = Array(2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2)
        arr2 = Array(2, 1, 0)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        arr = Array(2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        


        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(1, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 2
        col.Add 2
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1, 0)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(2, 1, 2)
        fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        



        Set col = New Collection
        col.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 1
        col2.Add 2
        col2.Add 3
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


        arr = Array(1)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        arr = Array(1)
        arr2 = Array(1)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


        arr = Array(1, 2)
        arr2 = Array(1, 2)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        arr = Array(1)
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(1, 2)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(1, 2, 3)
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 1
        col2.Add 1
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col2.Add 2
        col2.Add 2
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        Set col = New Collection
        Set col2 = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        col2.Add 3
        col2.Add 2
        col2.Add 1
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        arr = Array(1)
        arr2 = Array(2)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


        arr = Array(1, 2)
        arr2 = Array(2, 1)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(1, 2, 3)
        arr2 = Array(3, 2, 1)
        fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        arr = Array(2)
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        Set col = New Collection
        col.Add 1
        col.Add 2
        arr = Array(2, 1)
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 1
        col.Add 2
        col.Add 3
        arr = Array(3, 2, 1)
        fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    End With
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.Result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).Result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            resultBool = test.Result = .Tests(i).Result
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput
            
            Debug.Assert resultBool And fluentBool And valueBool And inputBool
            
            i = i + 1
        Next test
    End With
    
    Debug.Print "Negative tests finished"
    negTestCount = mTestCounter + mMiscNegTests
    mTestCounter = 0
    printTestCount (negTestCount)
    
    Set negativeDocumentationTests = testFluent
    
End Function

Public Function positiveNullDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    Dim col As Collection
    Dim d As Scripting.Dictionary
    Dim arr As Variant
    Dim test As ITest
    Dim nullBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim i As Long
        
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure("Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.EqualTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.EqualTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures("Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures("""Hello World""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(" ""Hello World"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures("Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures("""Hello World""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(" ""Hello World "" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Contain(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Contain(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.StartWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.StartWith(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.EndWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Have.LengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Have.MaxLengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Have.MinLengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello World!").Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""Hello World!""").Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(" ""Hello World!"" ").Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(""" Hello World! """).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(123).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(1.23).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(True).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello World!").Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""Hello World!""").Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(" ""Hello World!"" ").Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(""" Hello World! """).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(123).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(1.23).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(True).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(False).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(123).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(1.23).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(True).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(False).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Null).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(123).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(1.23).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(True).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(False).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Null).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.IdenticalTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.IdenticalTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.IdenticalTo(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.ExactSameElementsAs(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.ExactSameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.ExactSameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameUniqueElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameUniqueElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameUniqueElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameUniqueElementsAs(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameUniqueElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameUniqueElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameElementsAs(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    For Each test In fluent.Meta.Tests
        Debug.Assert test.HasNull
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).HasNull
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            nullBool = test.HasNull = .Tests(i).HasNull
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput

            Debug.Assert nullBool And fluentBool And valueBool And inputBool

            i = i + 1
        Next test
    End With
    
    Debug.Print "Positive null tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set positiveNullDocumentationTests = testFluent
End Function

Public Function negativeNullDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    Dim col As Collection
    Dim d As Scripting.Dictionary
    Dim arr As Variant
    Dim test As ITest
    Dim nullBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim i As Long
        
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure("Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.EqualTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.EqualTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures("Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures("""Hello World""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(" ""Hello World"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures("Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures("""Hello World""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(" ""Hello World "" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Contain(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.StartWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.EndWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.LengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.MaxLengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """Hello world!"""
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add " ""Hello world!"" "
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add """ Hello world! """
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.MinLengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello World!").ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""Hello World!""").ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(" ""Hello World!"" ").ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(""" Hello World! """).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(123).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(1.23).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(True).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello World!").ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""Hello World!""").ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(" ""Hello World!"" ").ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(""" Hello World! """).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(123).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(1.23).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(True).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(False).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(123).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(1.23).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(True).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(False).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(123).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(1.23).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(True).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(False).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.IdenticalTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.IdenticalTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.IdenticalTo(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.ExactSameElementsAs(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.ExactSameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.ExactSameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameUniqueElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameUniqueElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameUniqueElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameUniqueElementsAs(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameUniqueElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameUniqueElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.SameElementsAs(""" Hello world """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameElementsAs(Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    For Each test In fluent.Meta.Tests
        Debug.Assert test.HasNull
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).HasNull
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            nullBool = test.HasNull = .Tests(i).HasNull
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput

            Debug.Assert nullBool And fluentBool And valueBool And inputBool

            i = i + 1
        Next test
    End With
    
    Debug.Print "Negative null tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set negativeNullDocumentationTests = testFluent
End Function

Public Function emptyDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf) As IFluentOf
    Dim col As Collection
    Dim d As Scripting.Dictionary
    Dim arr As Variant
    Dim test As ITest
    Dim emptyBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim i As Long
        
    fluent.TestValue = testFluent.Of().Should.Be.EqualTo("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure("Hello World")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure("Hello World")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures("Hello World")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures("Hello World")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.EqualTo("""Hello world""")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo("""Hello world""")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure("""Hello world""")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure("""Hello world""")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures("""Hello world""")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures("""Hello world""")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    fluent.TestValue = testFluent.Of().Should.Be.EqualTo(" ""Hello world"" ")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo(" ""Hello world"" ")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure(" ""Hello world"" ")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure(" ""Hello world"" ")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures(" ""Hello world"" ")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures(" ""Hello world"" ")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.EqualTo(""" Hello world """)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo(""" Hello world """)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure(""" Hello world """)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure(""" Hello world """)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures(""" Hello world """)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures(""" Hello world """)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.GreaterThan(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.GreaterThan(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Be.GreaterThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.GreaterThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.LessThan(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.LessThan(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Be.LessThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.LessThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Contain("Hello world!")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Contain("Hello world!")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.StartWith("Hello")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.StartWith("Hello")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.EndWith("Hello")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.EndWith("Hello")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.LengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.LengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.MaxLengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.MaxLengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.MinLengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.MinLengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.Between(1, 10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.Between(1, 10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.LengthBetween(1, 10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().ShouldNot.Have.LengthBetween(1, 10)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.Alphabetic
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.Alphabetic
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Be.Numeric
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.Numeric
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Be.Alphanumeric
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.Alphanumeric
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Be.Erroneous
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.Erroneous
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().Should.Have.ErrorNumberOf("2007")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.ErrorNumberOf("2007")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.IdenticalTo("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.IdenticalTo("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.ExactSameElementsAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.ExactSameElementsAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.SameUniqueElementsAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.SameElementsAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.SameElementsAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.OneOf("Hello world", 5, True)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.OneOf("Hello world", 5, True)
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of().Should.Have.SameTypeAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.SameTypeAs("Hello world")
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    For Each test In testFluent.Meta.Tests
        Debug.Assert Not test.TestValueSet And test.HasEmpty
    Next test
    
    For i = 1 To testFluent.Meta.Tests.Count
        Debug.Assert Not testFluent.Meta.Tests(i).TestValueSet And testFluent.Meta.Tests(i).HasEmpty
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            emptyBool = test.HasEmpty = .Tests(i).HasEmpty And Not test.TestValueSet = Not .Tests(i).TestValueSet
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput

            Debug.Assert emptyBool And fluentBool And valueBool And inputBool

            i = i + 1
        Next test
    End With
    
    Debug.Print "Empty tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set emptyDocumentationTests = testFluent
End Function

Private Function MiscTests(fluent As IFluent)
    Dim testCount As Long
    Dim q As Object
    Dim elem As Variant
    
    Set mEvents.setFluentEventDuplicate = fluent

    'test to ensure fluent object's default TestValue value is equal to empty
    Debug.Assert Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
    
    'test to ensure that a duplicate test event is not raised since skipDupCheck
    'is set to true
    
    With fluent.Meta.Tests
        .SkipDupCheck = True
            Debug.Assert Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
        .SkipDupCheck = False
    End With
    
    'test to ensure that a duplicate test event is raised since skipDupCheck
    'is set to false
    
    Debug.Assert Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
    
    'test to ensure fluent object's TestValue property can return a value
    fluent.TestValue = fluent.TestValue
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    
    'test to ensure fluent object's TestValue property can return an object
    Set fluent.TestValue = New Collection
    Set fluent.TestValue = fluent.TestValue
    Debug.Assert fluent.Should.Be.Something
    
    'test to ensure that addDataStructure is working with non-default datastructure
    
    Set q = CreateObject("system.collections.Queue")
    
    q.Enqueue ("Hello")
    
    fluent.Meta.Tests.AddDataStructure q
    
    fluent.TestValue = "Hello"
    
    Debug.Assert fluent.Should.Be.InDataStructure(q)
    
    'test to ensure that StrTestValue and StrTestInput are working with non-default datastructure
    
    With fluent.Meta
        Debug.Assert .Tests(.Tests.Count).StrTestValue = "`Hello`"
        Debug.Assert .Tests(.Tests.Count).StrTestInput = "Queue(`Hello`)"
    End With
    
    'Clean string tests
    
    fluent.Meta.Tests.CleanTestValStr = True
    
    fluent.TestValue = """abc"""
    
    Debug.Assert fluent.Should.Be.EqualTo("abc")
    
    fluent.Meta.Tests.CleanTestValStr = False
    fluent.Meta.Tests.CleanTestInputStr = True
    
    fluent.TestValue = "abc"

    Debug.Assert fluent.Should.Be.EqualTo("""abc""")
    
    fluent.Meta.Tests.CleanTestValStr = True
    fluent.Meta.Tests.CleanTestInputStr = True

    fluent.TestValue = """abc"""
    
    fluent.Meta.Tests.CleanTestValStr = False
    fluent.Meta.Tests.CleanTestInputStr = False
    
    fluent.Meta.Tests.CleanTestStrings = True

    fluent.TestValue = """abc"""

    Debug.Assert fluent.Should.Be.EqualTo("""abc""")
    
    'Add to clean strings tests
    
    fluent.Meta.Tests.AddToCleanStringDict ("'")

    fluent.Meta.Tests.CleanTestValStr = True
    
    fluent.TestValue = "'abc def'"

    Debug.Assert fluent.Should.Be.EqualTo("abcdef")

    fluent.Meta.Tests.CleanTestValStr = False
    fluent.Meta.Tests.CleanTestInputStr = True
    
    fluent.TestValue = "abcdef"
    
    Debug.Assert fluent.Should.Be.EqualTo("'abc def'")

    fluent.Meta.Tests.CleanTestValStr = False
    fluent.Meta.Tests.CleanTestInputStr = True
    
    fluent.Meta.Tests.AddToCleanStringDict " ", "_", True
    
    fluent.Meta.Tests.CleanTestValStr = True
    fluent.Meta.Tests.CleanTestInputStr = False
    
    fluent.TestValue = "'abc def'"
    
    Debug.Assert fluent.Should.Be.EqualTo("abc_def")
    
    fluent.Meta.Tests.CleanTestValStr = False
    fluent.Meta.Tests.CleanTestInputStr = True
    
    fluent.TestValue = "abc_def"
    
    Debug.Assert fluent.Should.Be.EqualTo("'abc def'")
    
    'Explicit clean strings using cUtilities
    
    fluent.Meta.Tests.CleanTestStrings = False
    
    fluent.TestValue = """abc"""
    
    fluent.TestValue = fluent.Meta.Utilities.CleanString(fluent.TestValue)
    
    Debug.Assert fluent.Should.Be.EqualTo("abc")
    
    fluent.TestValue = "bcd"
    
    fluent.TestValue = fluent.Meta.Utilities.CleanString(fluent.TestValue)
    
    Debug.Assert fluent.Should.Be.EqualTo(fluent.Meta.Utilities.CleanString("""bcd"""))
    
    Debug.Print "Misc tests finished"
    
    testCount = fluent.Meta.Tests.Count
    printTestCount (testCount)
    
    MiscTests = testCount
End Function

Public Function checkResetCounters(fluent As IFluent, testFluent As IFluentOf)
    Dim b As Boolean
    
    testFluent.Meta.Tests.ResetCounter
    fluent.Meta.Tests.ResetCounter
    
    b = (testFluent.Meta.Tests.Count = 0 And fluent.Meta.Tests.Count = 0)
   
   checkResetCounters = b
End Function

Public Function getFluentCounts(fluent As IFluent)
    Dim test As ITest
    Dim d As Scripting.Dictionary
    Dim temp As String
    Dim elem As Variant
    Dim fn As String
    
    temp = ""
    Set d = New Scripting.Dictionary
    
    For Each test In fluent.Meta.Tests
        fn = test.functionName
        If Not d.Exists(fn) Then
            d.Add fn, 1
        Else
            d(fn) = d(fn) + 1
        End If
    Next test
    
    For Each elem In d.Keys
        temp = temp & elem & ": " & d(elem) & vbNewLine
    Next elem
    
    getFluentCounts = temp
End Function

Public Function getFluentOfCounts(fluentOf As IFluentOf)
    Dim test As ITest
    Dim d As Scripting.Dictionary
    Dim temp As String
    Dim elem As Variant
    Dim fn As String
    
    temp = ""
    Set d = New Scripting.Dictionary
    
    For Each test In fluentOf.Meta.Tests
        fn = test.functionName
        If Not d.Exists(fn) Then
            d.Add fn, 1
        Else
            d(fn) = d(fn) + 1
        End If
    Next test
    
    For Each elem In d.Keys
        temp = temp & elem & ": " & d(elem) & vbNewLine
    Next elem
    
    getFluentOfCounts = temp
End Function

Private Function validateNegativeCounters(testFluent As IFluentOf) As Boolean
    Dim d As Scripting.Dictionary
    Dim test As ITest
    Dim counter As Long
    Dim fn As String
    Dim elem As Variant
    Dim testDev As ITestDev
    
    Set d = New Scripting.Dictionary
    
    counter = 0
    
    For Each test In testFluent.Meta.Tests
        Set testDev = test
        If testDev.NegateValue Then
            fn = test.functionName
            If Not d.Exists(fn) Then
                d.Add fn, 1
            Else
                d(fn) = d(fn) + 1
            End If
        End If
    Next test
    
    For Each elem In d.Keys
        Debug.Assert d(elem) > 0
        counter = counter + 1
    Next elem
    
    validateNegativeCounters = (d.Count = counter)
End Function
