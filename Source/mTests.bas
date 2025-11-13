Attribute VB_Name = "mTests"
Option Explicit

Private Const G_TB_SKIP As Boolean = False

Private mCounter As Long
Private mTestCounter As Long
Private mMiscPosTests As Long
Private mMiscNegTests As Long
Private posTestCount As Long
Private negTestCount As Long
Private tfRecur As IFluentOf
Private tfIter As IFluentOf
Private mEvents As zEvents
Private mRecurIterFuncNamesDict As Scripting.Dictionary

Public Sub runMainTests()
    Dim fluent As IFluent
    Dim testFluent As IFluentOf
    Dim testFluentResult As IFluentOf
    Dim events As zEvents
    Dim nulTestFluent As IFluentOf
    Dim emptyTestFluent As IFluentOf
    Dim tempCounter As Long
    Dim TestingInfoDev As ITestingFunctionsInfoDev
    Dim elem As Variant
    Dim recurCount1 As Long, iterCount1 As Long
    Dim recurCount2 As Long, iterCount2 As Long
    
    Set mRecurIterFuncNamesDict = New Scripting.Dictionary
    Set fluent = MakeFluent
    Set testFluentResult = MakeFluentOf
    
    Set testFluent = getAndInitTestFluent
    Set mEvents = getAndInitEvent(fluent, testFluent, testFluentResult)
    
    Call runEqualPosNegTests(fluent, testFluent, testFluentResult)
    
'    testFluent.Meta.Printing.PrintToSheet

    Call runRecurIterTests(testFluent)
    
    tempCounter = mCounter

    Set fluent = MakeFluent
    Set testFluent = MakeFluentOf
    
    mCounter = 0

    mCounter = tempCounter

    mCounter = mCounter + cleanStringTests(fluent)
    
    Set fluent = MakeFluent

    mCounter = mCounter + MiscTests(fluent)
    
    Debug.Print "All tests Finished"
    
    Call printTestCount(mCounter + mMiscNegTests + mMiscPosTests)
    
    Call resetAndCheckCounters(mEvents, fluent, testFluent)
    
    'set local and module-level variables to default values
    tempCounter = 0
    
    Set mRecurIterFuncNamesDict = Nothing
    Set fluent = Nothing
    Set testFluentResult = Nothing
    Set testFluent = Nothing
    Set mEvents = Nothing
    Set tfIter = Nothing
    Set tfRecur = Nothing
End Sub

Private Function getAndInitTestFluent() As IFluentOf
    Dim testFluent As IFluentOf
    
    Set testFluent = MakeFluentOf
    
    With testFluent.Meta
        .Printing.PassedMessage = "Success"
        .Printing.FailedMessage = "Failure"
        .Printing.UnexpectedMessage = "What?"
        
        .tests.ToStrDev = True
    End With
    
    
    Set getAndInitTestFluent = testFluent
End Function

Private Function getAndInitEvent(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As zEvents
    
    Set mEvents = New zEvents
    
    Set mEvents.setFluent = fluent
    Set mEvents.setFluentOf = testFluent
    Set mEvents.setFluentEventOfResult = testFluentResult
    
    Set getAndInitEvent = mEvents
End Function

Private Sub runEqualPosNegTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf)
    Dim posTestFluent As IFluentOf
    Dim negTestFluent As IFluentOf
    Dim posAndNegTestFluent As IFluentOf
    Dim equalTestFluent As IFluentOf
    Dim equalDict As Scripting.Dictionary
    Dim posDict As Scripting.Dictionary
    Dim negDict As Scripting.Dictionary
    Dim elem As Variant
    Dim i As Long
    Dim equalTestingInfo As ITestingFunctionsInfos
    Dim posAndNegTestingInfo As ITestingFunctionsInfos
    Dim equalTestingInfoDict As Scripting.Dictionary
    Dim posAndNegTestingInfoDict As Scripting.Dictionary
    Dim counter As Long
    
    fluent.Meta.Printing.Category = "Fluent - EqualityTests"
    testFluent.Meta.Printing.Name = "Test Fluent - abc 123"
    testFluent.Meta.Printing.Category = "Test Fluent - EqualityTests"
    Set equalTestFluent = EqualityDocumentationTests(fluent, testFluent, testFluentResult)
    Set equalTestingInfo = equalTestFluent.Meta.tests.TestingFunctionsInfos
    Set equalTestingInfoDict = equalTestingInfo.TestFuncInfoToDict
    
    Set tfRecur = MakeFluentOf
    Set tfIter = MakeFluentOf
    
    tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    
    'The equality tests cannot use validateTestDict counter since it will fail since they
    'only run the equality tests
    
    testFluent.Meta.tests.resetTestingInfo

    fluent.Meta.Printing.Category = "Fluent - positiveAndNegativeDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - positiveAndNegativeDocumentationTests"

    Set posAndNegTestFluent = testFluent
    Set posAndNegTestFluent = EqualToTests(fluent, testFluent, testFluentResult)
    Set posAndNegTestFluent = GreaterThanTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = GreaterThanOrEqualToTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = LessThanTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = LessThanOrEqualToTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = ContainTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = StartWithTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = EndWithTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = LengthOfTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = MaxLengthOfTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = MinLengthOfTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = BetweenTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = LengthBetweenTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = OneOfTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = SomethingTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = EvaluateToTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = AlphabeticTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = NumericTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = AlphanumericTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = SameTypeAsTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = IdenticalToTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = ExactSameElementsAsTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = SameUniqueElementsAsTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = SameElementsAsTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = ProcedureTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = ElementsTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = ElementsInDataStructureTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = InDataStructureTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = InDataStructuresTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = DepthCountOfTests(fluent, posAndNegTestFluent, testFluentResult)
    Set posAndNegTestFluent = NestedCountOfTests(fluent, posAndNegTestFluent, testFluentResult)
    
    If Not G_TB_SKIP Then
        Set posAndNegTestFluent = ErroneousTests(fluent, posAndNegTestFluent, testFluentResult)
        Set posAndNegTestFluent = ErrorDescriptionOfTests(fluent, posAndNegTestFluent, testFluentResult)
        Set posAndNegTestFluent = ErrorNumberOfTests(fluent, posAndNegTestFluent, testFluentResult)

        counter = 0
    Else
        counter = 3
    End If
    
    Call CheckTestFuncInfos(posAndNegTestFluent)
    
    Set posAndNegTestingInfo = posAndNegTestFluent.Meta.tests.TestingFunctionsInfos
    Set posAndNegTestingInfoDict = posAndNegTestingInfo.TestFuncInfoToDict
    
    Debug.Assert posAndNegTestingInfo.validateTfiDictCounters(posAndNegTestingInfoDict, counter)

    Debug.Assert posTestCount = negTestCount
    
End Sub

Private Sub resetAndCheckCounters(ByVal events As zEvents, ByVal fluent As IFluent, ByVal testFluent As IFluentOf)
    mCounter = 0
    
    mTestCounter = 0
    
    mMiscNegTests = 0
    
    mMiscPosTests = 0

    Debug.Assert events.CheckTestCounters

    Debug.Assert checkResetCounters(fluent, testFluent)
End Sub

Private Sub printTestCount(ByVal testCount As Long)
    If testCount > 1 Then
        Debug.Print testCount & " tests finished!" & vbNewLine
    ElseIf testCount = 1 Then
        Debug.Print "1 Test finished!"
    End If
End Sub

Private Sub TrueAssertAndRaiseEvents(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf)
    Dim td As ITestDev
    Dim inputIter As String
    Dim inputRecur As String
    Dim valueIter As String
    Dim valueRecur As String

    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.tests.Count = mCounter

    With fluent.Meta.tests
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    If testFluent.Meta.tests.ToStrDev Then
        With testFluent.Meta.tests
            Set td = .item(.Count)
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

Private Sub FalseAssertAndRaiseEvents(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf)
    Dim td As ITestDev
    Dim inputIter As String
    Dim inputRecur As String
    Dim valueIter As String
    Dim valueRecur As String
    
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.tests.Count = mCounter

    With fluent.Meta.tests
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    If testFluent.Meta.tests.ToStrDev Then
        With testFluent.Meta.tests
            Set td = .item(.Count)
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

Private Sub NullAssertAndRaiseEvents(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf)
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.tests.Count = mCounter

    With fluent
        Debug.Assert testFluentResult.Of(.TestValue).Should.Be.EqualTo(Null)
        Debug.Assert testFluentResult.Of(.TestValue).ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With
End Sub

Private Sub EmptyAssertAndRaiseEvents(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf)
    
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
'    Debug.Assert testFluent.Meta.Tests(testFluent.Meta.Tests.Count).TestValueSet = False
    
    Debug.Assert testFluent.Meta.tests.Count = mCounter
    
    Debug.Assert testFluent.Meta.tests(mCounter).TestValueSet = False

    With fluent
        Debug.Assert VBA.Information.IsEmpty(fluent.TestValue)
    End With
End Sub

Private Function EqualityDocumentationTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim test As ITest
    Dim i As Long
    Dim resultBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim counter As Long
    
    counter = 0

    With fluent.Meta.tests
    
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
    
        testFluent.Meta.tests.ApproximateEqual = True
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
        
        testFluent.Meta.tests.ApproximateEqual = False
        
        '//Null and Empty tests
        
        fluent.TestValue = testFluent.Of(Null).Should.Be.EqualTo(Null)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.EqualTo(Null)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
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
    
    For Each test In fluent.Meta.tests
        Debug.Assert test.result
    Next test
    
    For i = 1 To fluent.Meta.tests.Count
        Debug.Assert fluent.Meta.tests(i).result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .tests
            resultBool = test.result = .tests(i).result
            fluentBool = test.FluentPath = .tests(i).FluentPath
            valueBool = test.testingValue = .tests(i).testingValue
            inputBool = test.testingInput = .tests(i).testingInput
            
            Debug.Assert resultBool And fluentBool And valueBool And inputBool
            
            i = i + 1
        Next test
    End With
    
    Debug.Print "Equality tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set EqualityDocumentationTests = testFluent
End Function

Sub validateTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf)
    Dim test As ITest
    Dim i As Long
    Dim resultBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim valueBool2 As Boolean
    Dim inputBool2 As Boolean
    Dim selfReferentialBool As Boolean
    Dim inputSelfReferential As Boolean
    Dim valueSelfReferential As Boolean
    
    For Each test In fluent.Meta.tests
        Debug.Assert test.result
    Next test
    
    For i = 1 To fluent.Meta.tests.Count
        Debug.Assert fluent.Meta.tests(i).result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .tests
            resultBool = test.result = .tests(i).result
            fluentBool = test.FluentPath = .tests(i).FluentPath
            valueBool = test.strTestValue = .tests(i).strTestValue
            inputBool = test.StrTestInput = .tests(i).StrTestInput
            valueBool2 = test.strTestValuePretty = .tests(i).strTestValuePretty
            inputBool2 = test.StrTestInputPretty = .tests(i).StrTestInputPretty
            
            If test.HasSelfReferential And .tests(i).HasSelfReferential Then
                If VBA.Information.IsNull(test.TestingInputIsSelfReferential) Then
                    inputSelfReferential = VBA.Information.IsNull(test.TestingInputIsSelfReferential) And VBA.Information.IsNull(.tests(i).TestingInputIsSelfReferential)
                Else
                    inputSelfReferential = test.TestingInputIsSelfReferential = .tests(i).TestingInputIsSelfReferential
                End If
                
                If VBA.Information.IsNull(test.TestingValueIsSelfReferential) Then
                    valueSelfReferential = VBA.Information.IsNull(test.TestingValueIsSelfReferential) And VBA.Information.IsNull(.tests(i).TestingValueIsSelfReferential)
                Else
                    valueSelfReferential = test.TestingValueIsSelfReferential = .tests(i).TestingValueIsSelfReferential
                End If
                
                selfReferentialBool = inputSelfReferential And valueSelfReferential
                
                Debug.Assert resultBool And fluentBool And valueBool And inputBool And selfReferentialBool
            Else
                Debug.Assert resultBool And fluentBool And valueBool And inputBool
            End If
            
            i = i + 1
        Next test
    End With
End Sub

Sub validateRecurIterFluentOfs(ByVal testFluent As cFluentOf, ByVal tfRecur As cFluentOf, ByVal tfIter As cFluentOf, ByVal recurIterFuncName As String)
    Dim test As ITestDev
    Dim test2 As cTest
    Dim implicitRecurCount As Long
    Dim explicitRecurCount As Long
    Dim explicitIterCount As Long
    Dim b1 As Boolean
    Dim b2 As Boolean
    
    implicitRecurCount = 0
    explicitRecurCount = 0
    explicitIterCount = 0
    b1 = False
    b2 = False
    
    For Each test In testFluent.Meta.tests
        If Not VBA.Information.IsNull(test.Algorithm) Then
            Set test2 = test
            
            b1 = test.AlgorithmValueSet = False 'False because testFluent should use implicit flAlgorithm.flRecursive
            b2 = test.Algorithm = flAlgorithm.flRecursive
            
            Debug.Assert b1
            Debug.Assert b2
            Debug.Assert test.IsRecurIterFunc
            
            If b1 And b2 Then
                implicitRecurCount = implicitRecurCount + 1
            End If
        End If
    Next test
    
    b1 = False
    b2 = False
    
    For Each test In tfRecur.Meta.tests
        b1 = test.AlgorithmValueSet = True 'True because tfRecur should use explicit flAlgorithm.flRecursive
        b2 = test.Algorithm = flAlgorithm.flRecursive
        
        Debug.Assert b1
        Debug.Assert b2
        Debug.Assert test.IsRecurIterFunc
        Debug.Assert test.IsBaseCaseRecur
        
        If b1 And b2 Then
            explicitRecurCount = explicitRecurCount + 1
        End If
    Next test
    
    b1 = False
    b2 = False
    
    For Each test In tfIter.Meta.tests
        b1 = test.AlgorithmValueSet = True 'True because tfIter should use explicit flAlgorithm.flIterative
        b2 = test.Algorithm = flAlgorithm.flIterative
        
        Debug.Assert b1
        Debug.Assert b2
        Debug.Assert test.IsRecurIterFunc
        Debug.Assert test.IsBaseCaseIter
        
        If b1 And b2 Then
            explicitIterCount = explicitIterCount + 1
        End If
    Next test
    
    Debug.Assert (implicitRecurCount = explicitRecurCount) And (explicitRecurCount = explicitIterCount)
    
    If recurIterFuncName <> "main" Then
        mRecurIterFuncNamesDict.Add recurIterFuncName, recurIterFuncName
    End If
End Sub

Function validateRecurIterFuncCounts(ByVal recurIterFluentOf As cFluentOf) As Long
    Dim TestingInfoDev As ITestingFunctionsInfoDev
    Dim counter As Long
    
    With recurIterFluentOf.Meta.tests
        Set TestingInfoDev = .TestingFunctionsInfos
        
        With TestingInfoDev
            Debug.Assert .DepthCountOfIter.Count = .DepthCountOfRecur.Count
            If .DepthCountOfIter.Count = .DepthCountOfRecur.Count Then counter = counter + 1
            
            Debug.Assert .InDataStructureIter.Count = .InDataStructureRecur.Count
            If .DepthCountOfIter.Count = .DepthCountOfRecur.Count Then counter = counter + 1
            
            Debug.Assert .InDataStructuresIter.Count = .InDataStructuresRecur.Count
            If .DepthCountOfIter.Count = .DepthCountOfRecur.Count Then counter = counter + 1
            
            Debug.Assert .NestedCountOfIter.Count = .NestedCountOfRecur.Count
            If .DepthCountOfIter.Count = .DepthCountOfRecur.Count Then counter = counter + 1
        End With
    End With
    
    validateRecurIterFuncCounts = counter
End Function

Function validateRecurIterFuncCounts2(ByVal recurIterFluentOf As cFluentOf) As Long
    Dim TestingInfoDev As ITestingFunctionsInfoDev
    Dim counter As Long
    Dim recurIterFuncNameCol As VBA.Collection
    Dim elem As Variant
    Dim testSubInfoRecur As ITestingFunctionsInfo
    Dim testSubInfoIter As ITestingFunctionsInfo
    
    Set TestingInfoDev = recurIterFluentOf.Meta.tests.TestingFunctionsInfos
    Set recurIterFuncNameCol = TestingInfoDev.getRecurIterFuncNameCol
    counter = 0
    
    For Each elem In recurIterFuncNameCol
        Set testSubInfoRecur = VBA.Interaction.CallByName(TestingInfoDev, elem & "Recur", VbGet)
        Set testSubInfoIter = VBA.Interaction.CallByName(TestingInfoDev, elem & "Iter", VbGet)
        
        Debug.Assert testSubInfoIter.Count = testSubInfoRecur.Count
        
        If testSubInfoIter.Count = testSubInfoRecur.Count Then counter = counter + 1
    Next elem
    
    validateRecurIterFuncCounts2 = counter
End Function

Function validateRecurIterFuncNamesFromFluentOfInDict(ByVal recurIterFluentOf As cFluentOf) As Boolean
    Dim elem As Variant
    Dim counter As Long
    Dim recurIterFuncNamesCol As VBA.Collection
    Dim TestingInfoDev As ITestingFunctionsInfoDev
    
    Set TestingInfoDev = recurIterFluentOf.Meta.tests.TestingFunctionsInfos
    Set recurIterFuncNamesCol = TestingInfoDev.getRecurIterFuncNameCol
    
    Debug.Assert recurIterFuncNamesCol.Count = mRecurIterFuncNamesDict.Count
    
    For Each elem In recurIterFuncNamesCol
        If mRecurIterFuncNamesDict.Exists(elem) Then
            counter = counter + 1
        End If
    Next elem
    
    validateRecurIterFuncNamesFromFluentOfInDict = (recurIterFuncNamesCol.Count) = counter And (mRecurIterFuncNamesDict.Count = counter)
End Function

Sub runRecurIterTests(ByVal testFluent As cFluentOf)
    Dim recurCount1 As Long, iterCount1 As Long
    Dim recurCount2 As Long, iterCount2 As Long
    
    Call validateRecurIterFluentOfs(testFluent, tfRecur, tfIter, "main")
    
    Call validateRecurIterFuncCounts2(tfRecur)
    
    recurCount1 = validateRecurIterFuncCounts(tfRecur)
    iterCount1 = validateRecurIterFuncCounts(tfIter)
    Debug.Assert recurCount1 = iterCount1
    
    recurCount2 = validateRecurIterFuncCounts2(tfRecur)
    iterCount2 = validateRecurIterFuncCounts2(tfIter)
    Debug.Assert recurCount2 = iterCount2
    
    Debug.Assert _
    (recurCount1 = iterCount1) And _
    (recurCount2 = iterCount2) And _
    (recurCount1 = recurCount2) And _
    (iterCount1 = iterCount2)
    
    Debug.Assert _
    validateRecurIterFuncNamesFromFluentOfInDict(tfRecur) And _
    validateRecurIterFuncNamesFromFluentOfInDict(tfIter)
End Sub

Private Function EqualToTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    fluent.TestValue = testFluent.Of("""abc""").Should.Be.EqualTo("""abc""")
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.EqualTo(10)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    ' //Approximate equality tests
    testFluent.Meta.tests.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("10").Should.Be.EqualTo(10)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("True").Should.Be.EqualTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    fluent.TestValue = testFluent.Of(5.0000001).Should.Be.EqualTo(5)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(CStr(Excel.Evaluate("1 / 0"))).Should.Be.EqualTo("Error 2007")
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""abc""").ShouldNot.Be.EqualTo("""abc""")
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.EqualTo(10)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(CStr(Excel.Evaluate("1 / 0"))).ShouldNot.Be.EqualTo("Error 2007")
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    ' //Approximate equality tests
    testFluent.Meta.tests.ApproximateEqual = True
    
    fluent.TestValue = testFluent.Of("10").ShouldNot.Be.EqualTo(10)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("True").ShouldNot.Be.EqualTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    fluent.TestValue = testFluent.Of(5.0000001).ShouldNot.Be.EqualTo(5)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.EqualTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.EqualTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.EqualTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.EqualTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.EqualTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty tests

    fluent.TestValue = testFluent.Of().Should.Be.EqualTo("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().Should.Be.EqualTo("""Hello world""")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo("""Hello world""")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().Should.Be.EqualTo(" ""Hello world"" ")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo(" ""Hello world"" ")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().Should.Be.EqualTo(""" Hello world """)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo(""" Hello world """)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Call validateTests(fluent, testFluent)
    
    Debug.Print "EqualToTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set EqualToTests = testFluent

End Function

Private Function GreaterThanTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
        
    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(9)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(11)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(9)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(11)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null tests
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null tests
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documentation tests
    
    fluent.TestValue = testFluent.Of().Should.Be.GreaterThan(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.GreaterThan(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "GreaterThanTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set GreaterThanTests = testFluent
End Function

Private Function LessThanTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(9)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(11)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(9)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(11)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'negative null documentation tests
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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

    'empty documentation tests
    fluent.TestValue = testFluent.Of().Should.Be.LessThan(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.LessThan(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Call validateTests(fluent, testFluent)
    
    Debug.Print "LessThanTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set LessThanTests = testFluent
End Function

Private Function ContainTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
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

    'negative documentation tests
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

    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Contain(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Contain(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Contain(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain("""Hello world!""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain(" ""Hello world!"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain(""" Hello world! """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Contain(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'empty tests
    fluent.TestValue = testFluent.Of().Should.Contain("Hello world!")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Contain("Hello world!")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Call validateTests(fluent, testFluent)
    
    Debug.Print "ContainTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ContainTests = testFluent
End Function

Private Function StartWithTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
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

    'negative documentation tests
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

    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.StartWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.StartWith(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.StartWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.StartWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documentation tests
    fluent.TestValue = testFluent.Of().Should.StartWith("Hello")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.StartWith("Hello")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "StartWithTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set StartWithTests = testFluent
End Function

Private Function EndWithTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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

    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.EndWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.EndWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("""Hello""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith(" ""Hello"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.EndWith(""" Hello """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'empty documention tests
    fluent.TestValue = testFluent.Of().Should.EndWith("Hello")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.EndWith("Hello")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "EndWithTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set EndWithTests = testFluent
End Function

Private Function LengthOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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

    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.LengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.LengthOf(2)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.LengthOf(2)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "LengthOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set LengthOfTests = testFluent
End Function

Private Function MaxLengthOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.MaxLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.MaxLengthOf(2)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.MaxLengthOf(2)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "MaxLengthOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set MaxLengthOfTests = testFluent
End Function

Private Function MinLengthOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'empty documention tests

    fluent.TestValue = testFluent.Of().Should.Have.MinLengthOf(2)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.MinLengthOf(2)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "MinLengthOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set MinLengthOfTests = testFluent
End Function

Private Function GreaterThanOrEqualToTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Be.GreaterThanOrEqualTo(10)
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Be.GreaterThanOrEqualTo(10)
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.GreaterThanOrEqualTo(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.GreaterThanOrEqualTo(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "GreaterThanOrEqualToTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set GreaterThanOrEqualToTests = testFluent
End Function

Private Function LessThanOrEqualToTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array()).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
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
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.LessThanOrEqualTo(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.LessThanOrEqualTo(10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "LessThanOrEqualToTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set LessThanOrEqualToTests = testFluent
End Function

Private Function BetweenTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of("Hello World!").Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""Hello World!""").Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(" ""Hello World!"" ").Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(""" Hello World! """).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of("Hello World!").ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("""Hello World!""").ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(" ""Hello World!"" ").ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(""" Hello World! """).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Between(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.Between(1, 10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.Between(1, 10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "BetweenTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set BetweenTests = testFluent
End Function

Private Function LengthBetweenTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    'negative documentation tests
    
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
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.LengthBetween(1, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.LengthBetween(1, 10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.LengthBetween(1, 10)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "LengthBetweenTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set LengthBetweenTests = testFluent
End Function

Private Function OneOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
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
    
    Set col = New VBA.Collection
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(col).Should.Be.OneOf(col, d)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    Set d = Nothing
    
    fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(col, d, 10)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative documentation tests
    
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
    
    Set col = New VBA.Collection
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.OneOf(col, d)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    Set d = Nothing
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(col, d, 10)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'OneOf does not have positive or negative null documentation tests
        
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.OneOf("Hello world", 5, True)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.OneOf("Hello world", 5, True)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "OneOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set OneOfTests = testFluent
End Function

Private Function SomethingTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Something
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = Nothing
    fluent.TestValue = testFluent.Of(col).Should.Be.Something
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative documentation tests
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Something
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = Nothing
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Something
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
                        
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Something
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.Something
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.Something
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "SomethingTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set SomethingTests = testFluent
End Function

Private Function InDataStructureTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim col2 As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim arr() As Variant
    Dim strArr(1, 1) As Variant
    Dim b As Boolean
    Dim al As Object
    Dim val As Variant
    Dim tfBitwiseFlag As cFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev
    
'positive documentation tests
    
    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(False) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(9, 10, 11)
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
    
    
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add 10
    col.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(col) = tfIter.Of(10).Should.Be.InDataStructure(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
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
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
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
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(al) = tfIter.Of(10).Should.Be.InDataStructure(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set d = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary

    strArr(0, 0) = "Hello"
    strArr(1, 0) = "World"
    strArr(0, 1) = "Goodbye"
    strArr(1, 1) = "World"

    col.Add 1
    col2.Add 2
    col2.Add VBA.[_HiddenModule].Array(4, strArr)
    col.Add col2
    col.Add 6
    d2.Add "B", col
    d.Add "A", d2
    
    fluent.TestValue = testFluent.Of(1).Should.Be.InDataStructure(d) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(1).Should.Be.InDataStructure(d) = tfIter.Of(1).Should.Be.InDataStructure(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'negative documentation tests
    
    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(False) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
    arr = VBA.[_HiddenModule].Array(9, 10, 11)
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
    
    
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add 10
    col.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(col) = tfIter.Of(10).ShouldNot.Be.InDataStructure(col)
    mMiscNegTests = mMiscNegTests + 1
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
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
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
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
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(al) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(al) = tfIter.Of(10).ShouldNot.Be.InDataStructure(al)
    mMiscNegTests = mMiscNegTests + 1
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set d = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary

    strArr(0, 0) = "Hello"
    strArr(1, 0) = "World"
    strArr(0, 1) = "Goodbye"
    strArr(1, 1) = "World"

    col.Add 1
    col2.Add 2
    col2.Add VBA.[_HiddenModule].Array(4, strArr)
    col.Add col2
    col.Add 6
    d2.Add "B", col
    d.Add "A", d2
    
    fluent.TestValue = testFluent.Of(1).ShouldNot.Be.InDataStructure(d) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(1).ShouldNot.Be.InDataStructure(d) = tfIter.Of(1).ShouldNot.Be.InDataStructure(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'positive null documentation tests

    val = "Hello World"
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """ Hello World """
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 10
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 123.45
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = True
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = Null
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'negative null documentation tests
    
    val = "Hello World"
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = """ Hello World """
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(""" Hello World """)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = 10
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 123.45
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = True
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = Null
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'empty documention tests
    
    val = "Hello World"
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """Hello world"""
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = " ""Hello world"" "
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure(" ""Hello world"" ")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """ Hello world """
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructure(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    
    Call validateRecurIterFluentOfs(testFluent, tfRecur, tfIter, "InDataStructure")
    
    Call validateTests(fluent, testFluent)
    
'bitwise flags tests

    Set tfBitwiseFlag = MakeFluentOf
    tfBitwiseFlag.Meta.tests.Algorithm = flAlgorithm.flIterative + flAlgorithm.flRecursive
    
    arr = VBA.[_HiddenModule].Array(9, 10, 11)
    Debug.Assert tfBitwiseFlag.Of(10).Should.Be.InDataStructure(arr) 'with implicit recur
    
    arr = VBA.[_HiddenModule].Array(9, 11)
    Debug.Assert tfBitwiseFlag.Of(10).ShouldNot.Be.InDataStructure(arr) 'with implicit recur
    
    Set testInfoDev = tfBitwiseFlag.Meta.tests.TestingFunctionsInfos
    
    With testInfoDev
        Debug.Assert .InDataStructureRecur.Count > 0 And .InDataStructureIter.Count > 0
        Debug.Assert .InDataStructureRecur.Count = .InDataStructureIter.Count
    End With
    
    Set tfBitwiseFlag = Nothing
    
    Debug.Print "InDataStructureTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set InDataStructureTests = testFluent
End Function

Private Function InDataStructuresTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim b As Boolean
    Dim al As Object
    Dim val As Variant
    Dim tfBitwiseFlag As cFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev

'positive documentation tests
           
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(arr2) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(arr2) = tfIter.Of(10).Should.Be.InDataStructures(arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    fluent.TestValue = testFluent.Of(12).Should.Be.InDataStructures(arr, arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(12).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(12).Should.Be.InDataStructures(arr, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    fluent.TestValue = testFluent.Of(9).Should.Be.InDataStructures(arr, arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(9).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(9).Should.Be.InDataStructures(arr, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(arr, arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(10).Should.Be.InDataStructures(arr, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    fluent.TestValue = testFluent.Of(13).Should.Be.InDataStructures(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(13).Should.Be.InDataStructures(col) = tfIter.Of(13).Should.Be.InDataStructures(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    fluent.TestValue = testFluent.Of(16).Should.Be.InDataStructures(arr, col, arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(16).Should.Be.InDataStructures(arr, col, arr2) = tfIter.Of(16).Should.Be.InDataStructures(arr, col, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
                    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(col) = tfIter.Of(10).Should.Be.InDataStructures(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    
    arr = VBA.[_HiddenModule].Array(12, 13, 14)
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    fluent.TestValue = testFluent.Of(14).Should.Be.InDataStructures(col, arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(14).Should.Be.InDataStructures(col, arr) = tfIter.Of(14).Should.Be.InDataStructures(col, arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(d)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(d) = tfIter.Of(10).Should.Be.InDataStructures(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(6, VBA.[_HiddenModule].Array(7, VBA.[_HiddenModule].Array(8)))
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(8).Should.Be.InDataStructures(al, arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(8).Should.Be.InDataStructures(al, arr) = tfIter.Of(8).Should.Be.InDataStructures(al, arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'negative documentation tests
    
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(arr2) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(arr2) = tfIter.Of(10).ShouldNot.Be.InDataStructures(arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    fluent.TestValue = testFluent.Of(12).ShouldNot.Be.InDataStructures(arr, arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(12).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(12).ShouldNot.Be.InDataStructures(arr, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    fluent.TestValue = testFluent.Of(9).ShouldNot.Be.InDataStructures(arr, arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(9).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(9).ShouldNot.Be.InDataStructures(arr, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(arr, arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(10).ShouldNot.Be.InDataStructures(arr, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    fluent.TestValue = testFluent.Of(13).ShouldNot.Be.InDataStructures(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(13).ShouldNot.Be.InDataStructures(col) = tfIter.Of(13).ShouldNot.Be.InDataStructures(col)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''
    
    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    fluent.TestValue = testFluent.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2) = tfIter.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(col) = tfIter.Of(10).ShouldNot.Be.InDataStructures(col)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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

    arr = VBA.[_HiddenModule].Array(12, 13, 14)
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    fluent.TestValue = testFluent.Of(14).ShouldNot.Be.InDataStructures(col, arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(14).ShouldNot.Be.InDataStructures(col, arr) = tfIter.Of(14).ShouldNot.Be.InDataStructures(col, arr)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(d)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(d) = tfIter.Of(10).ShouldNot.Be.InDataStructures(d)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
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

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(al)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    arr = VBA.[_HiddenModule].Array(6, VBA.[_HiddenModule].Array(7, VBA.[_HiddenModule].Array(8)))
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(8).ShouldNot.Be.InDataStructures(al, arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(8).ShouldNot.Be.InDataStructures(al, arr) = tfIter.Of(8).ShouldNot.Be.InDataStructures(al, arr)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(al)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

'positive null documentation tests
    
    val = "Hello World"
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """Hello World"""
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures("""Hello World""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = " ""Hello World"" "
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """ Hello World """
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = 10
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = 123.45
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = True
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = Null
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = "Hello World"
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """Hello World"""
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = " ""Hello World "" "
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """ Hello World """
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 10
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 123.45
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = True
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = Null
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'negative null documentation tests
    
    val = "Hello World"
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """Hello World"""
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures("""Hello World""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = " ""Hello World"" "
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """ Hello World """
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = 10
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = 123.45
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = True
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = Null
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = "Hello World"
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """Hello World"""
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = " ""Hello World "" "
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = """ Hello World """
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 10
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 123.45
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = True
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = Null
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'empty documention tests
    
    val = "Hello World"
    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures("""Hello world""")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures("""Hello world""")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures(" ""Hello world"" ")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures(" ""Hello world"" ")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().Should.Be.InDataStructures(""" Hello world """)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.InDataStructures(""" Hello world """)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive

    Call validateRecurIterFluentOfs(testFluent, tfRecur, tfIter, "InDataStructures")
    
    Call validateTests(fluent, testFluent)
    
'bitwise flags tests

    Set tfBitwiseFlag = MakeFluentOf
    tfBitwiseFlag.Meta.tests.Algorithm = flAlgorithm.flIterative + flAlgorithm.flRecursive
    
    arr = VBA.[_HiddenModule].Array(9, 10, 11)
    Debug.Assert tfBitwiseFlag.Of(10).Should.Be.InDataStructures(arr) 'with implicit recur
    
    arr = VBA.[_HiddenModule].Array(9, 11)
    Debug.Assert tfBitwiseFlag.Of(10).ShouldNot.Be.InDataStructures(arr) 'with implicit recur
    
    Set testInfoDev = tfBitwiseFlag.Meta.tests.TestingFunctionsInfos
    
    With testInfoDev
        Debug.Assert .InDataStructuresRecur.Count > 0 And .InDataStructuresIter.Count > 0
        Debug.Assert .InDataStructuresRecur.Count = .InDataStructuresIter.Count
    End With
    
    Set tfBitwiseFlag = Nothing
    
    Debug.Print "InDataStructuresTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set InDataStructuresTests = testFluent
End Function



Private Function EvaluateToTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim arr() As Variant

    'positive documentation tests
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
    
    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(VBA.Information.TypeName(arr) = "Variant()").Should.EvaluateTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.Information.IsArray(arr)).Should.EvaluateTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(VBA.Information.TypeName(col) = "Collection").Should.EvaluateTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(TypeOf col Is Collection).Should.EvaluateTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(VBA.Information.TypeName(d) = "Dictionary").Should.EvaluateTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(TypeOf d Is Scripting.Dictionary).Should.EvaluateTo(True)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    '//Testing errors is possible if they're put in strings
    fluent.TestValue = testFluent.Of("1 / 0").Should.EvaluateTo(CVErr(xlErrDiv0))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'negative documentation tests
                
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
    
    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(VBA.Information.TypeName(arr) = "Variant()").ShouldNot.EvaluateTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(VBA.Information.IsArray(arr)).ShouldNot.EvaluateTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(VBA.Information.TypeName(col) = "Collection").ShouldNot.EvaluateTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(TypeOf col Is Collection).ShouldNot.EvaluateTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(VBA.Information.TypeName(d) = "Dictionary").ShouldNot.EvaluateTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(TypeOf d Is Scripting.Dictionary).ShouldNot.EvaluateTo(True)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    '//Testing errors is possible if they're put in strings
    fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.EvaluateTo(CVErr(xlErrDiv0))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.EvaluateTo(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.EvaluateTo(True)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.EvaluateTo(True)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "EvaluateToTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set EvaluateToTests = testFluent
End Function




Private Function AlphabeticTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
        
    fluent.TestValue = testFluent.Of("abc").Should.Be.Alphabetic
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("abc!@#").Should.Be.Alphabetic
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("123").Should.Be.Alphabetic
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("!@#").Should.Be.Alphabetic
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = True
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = False

    'negative documentation tests
    
    fluent.TestValue = testFluent.Of("abc").ShouldNot.Be.Alphabetic
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("abc!@#").ShouldNot.Be.Alphabetic
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("123").ShouldNot.Be.Alphabetic
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("!@#").ShouldNot.Be.Alphabetic
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = True
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = False
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Alphabetic
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.Alphabetic
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.Alphabetic
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "AlphabeticTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set AlphabeticTests = testFluent
End Function

Private Function NumericTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = True
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = False
    
    'negative documentation tests
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = True
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = False
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Numeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.Numeric
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.Numeric
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "NumericTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set NumericTests = testFluent
End Function

Private Function AlphanumericTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
    
    fluent.TestValue = testFluent.Of("abc123").Should.Be.Alphanumeric
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("abc").Should.Be.Alphanumeric
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("123").Should.Be.Alphanumeric
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("!@#").Should.Be.Alphanumeric
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = True
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = False
    
    'negative documentation tests
    
    fluent.TestValue = testFluent.Of("abc123").ShouldNot.Be.Alphanumeric
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("abc").ShouldNot.Be.Alphanumeric
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("123").ShouldNot.Be.Alphanumeric
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("!@#").ShouldNot.Be.Alphanumeric
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = True
    
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
    
    testFluent.Meta.tests.TestStrings.CleanTestStrings = False

    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.Alphanumeric
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.Alphanumeric
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.Alphanumeric
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "AlphanumericTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set AlphanumericTests = testFluent
End Function

Private Function ErroneousTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
    
    fluent.TestValue = testFluent.Of("1 / 0").Should.Be.Erroneous
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        fluent.TestValue = testFluent.Of(Err).Should.Be.Erroneous
    On Error GoTo 0
    
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) 'good
                
    'negative documentation tests
    
    fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.Be.Erroneous
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        fluent.TestValue = testFluent.Of(Err).ShouldNot.Be.Erroneous
    On Error GoTo 0
    
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
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
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Be.Erroneous
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.Erroneous
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.Erroneous
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ErroneousTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ErroneousTests = testFluent
End Function

Private Function ErrorNumberOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
    
    fluent.TestValue = testFluent.Of("1 / 0").Should.Have.ErrorNumberOf(2007)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        fluent.TestValue = testFluent.Of(Err).Should.Have.ErrorNumberOf(11)
    On Error GoTo 0
    
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative documentation tests
    
    fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.Have.ErrorNumberOf(2007)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        fluent.TestValue = testFluent.Of(Err).ShouldNot.Have.ErrorNumberOf(11)
    On Error GoTo 0
    
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    
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

    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
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

    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.ErrorNumberOf("2007")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.ErrorNumberOf("2007")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.ErrorNumberOf("2007")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ErrorNumberOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ErrorNumberOfTests = testFluent
End Function

Private Function SameTypeAsTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim arr() As Variant

    'positive documentation tests
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
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameTypeAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Nothing).Should.Have.SameTypeAs(Nothing)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(CLng(123)).Should.Have.SameTypeAs(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    Set col = Nothing

    'negative documentation tests
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
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameTypeAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Nothing).ShouldNot.Have.SameTypeAs(Nothing)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(CLng(123)).ShouldNot.Have.SameTypeAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    Set col = Nothing
    
    'SameTypeAs does not have positive or negative null documentation tests
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.SameTypeAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.SameTypeAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "SameTypeAsTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set SameTypeAsTests = testFluent
End Function

Private Function IdenticalToTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
    With testFluent.Of(col).Should.Be
        fluent.TestValue = .IdenticalTo(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1
    
    With testFluent.Of(col).Should.Be
        fluent.TestValue = .IdenticalTo(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1
    
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    fluent.TestValue = testFluent.Of(col2).Should.Be.IdenticalTo(col3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(col3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    fluent.TestValue = testFluent.Of(col2).Should.Be.IdenticalTo(col3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col2).Should.Be.IdenticalTo(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
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
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).Should.Be.IdenticalTo(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    fluent.TestValue = testFluent.Of(arr).Should.Be.IdenticalTo(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
    With testFluent.Of(arr).Should.Be
        fluent.TestValue = .IdenticalTo(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(VBA.[_HiddenModule].Array(1, 2, 3))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    fluent.TestValue = testFluent.Of(arr).Should.Be.IdenticalTo(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).Should.Be
        fluent.TestValue = .IdenticalTo(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
    With testFluent.Of(col).ShouldNot.Be
        fluent.TestValue = .IdenticalTo(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1
    
    With testFluent.Of(col).ShouldNot.Be
        fluent.TestValue = .IdenticalTo(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    fluent.TestValue = testFluent.Of(col2).ShouldNot.Be.IdenticalTo(col3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(col3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    fluent.TestValue = testFluent.Of(col2).ShouldNot.Be.IdenticalTo(col3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col2).ShouldNot.Be.IdenticalTo(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
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
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Be.IdenticalTo(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Be.IdenticalTo(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
    With testFluent.Of(arr).ShouldNot.Be
        fluent.TestValue = .IdenticalTo(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(VBA.[_HiddenModule].Array(1, 2, 3))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Be.IdenticalTo(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).ShouldNot.Be
        fluent.TestValue = .IdenticalTo(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .IdenticalTo(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    'positive null documentation tests
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.IdenticalTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Be.IdenticalTo(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.IdenticalTo(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Be.IdenticalTo(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Be.IdenticalTo("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Be.IdenticalTo("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "IdenticalToTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set IdenticalToTests = testFluent
End Function

Private Function ExactSameElementsAsTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
    With testFluent.Of(col).Should.Have
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1

    With testFluent.Of(col).Should.Have
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1

    With testFluent.Of(col2).Should.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2

    With testFluent.Of(col3).Should.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col2).Should.Have.ExactSameElementsAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
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
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).Should.Have.ExactSameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    fluent.TestValue = testFluent.Of(arr).Should.Have.ExactSameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

    With testFluent.Of(arr).Should.Have
        fluent.TestValue = .ExactSameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    fluent.TestValue = testFluent.Of(arr).Should.Have.ExactSameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).Should.Have
        fluent.TestValue = .ExactSameElementsAs(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs(VBA.[_HiddenModule].Array(2))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1

    With testFluent.Of(VBA.[_HiddenModule].Array(1)).Should.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1

    With testFluent.Of(col).Should.Have
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2

    With testFluent.Of(col2).Should.Have
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2

    With testFluent.Of(col2).Should.Have
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col2).Should.Have.ExactSameElementsAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    

    Set col = New VBA.Collection
    col.Add 1
    With testFluent.Of(VBA.[_HiddenModule].Array(2)).Should.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    Set col = Nothing

    
    'negative documentation tests
    
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
    With testFluent.Of(col).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1

    With testFluent.Of(col).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1

    With testFluent.Of(col2).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2

    With testFluent.Of(col3).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col2).ShouldNot.Have.ExactSameElementsAs(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
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
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.ExactSameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.ExactSameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

    With testFluent.Of(arr).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.ExactSameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs(VBA.[_HiddenModule].Array(2))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1

    With testFluent.Of(VBA.[_HiddenModule].Array(1)).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1

    With testFluent.Of(col).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2

    With testFluent.Of(col2).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2

    With testFluent.Of(col2).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col2).ShouldNot.Have.ExactSameElementsAs(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    With testFluent.Of(VBA.[_HiddenModule].Array(2)).ShouldNot.Have
        fluent.TestValue = .ExactSameElementsAs(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        fluent.TestValue = .ExactSameElementsAs(VBA.[_HiddenModule].Array(1))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    End With
    Set col = Nothing
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.ExactSameElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.ExactSameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.ExactSameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.ExactSameElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.ExactSameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.ExactSameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.ExactSameElementsAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.ExactSameElementsAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ExactSameElementsAsTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ExactSameElementsAsTests = testFluent
End Function

Private Function SameUniqueElementsAsTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant
    
    'positive documentation tests
    
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    col.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 3
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    col.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 0)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 2
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 0)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative documentation tests
    
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    col.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 3
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    col.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 0)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameUniqueElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    


    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 2
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 0)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameUniqueElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameUniqueElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameUniqueElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameUniqueElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameUniqueElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameUniqueElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameUniqueElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameUniqueElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.SameUniqueElementsAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.SameUniqueElementsAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "SameUniqueElementsAsTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set SameUniqueElementsAsTests = testFluent
End Function

Private Function SameElementsAsTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 1
    col2.Add 2
    col2.Add 3
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 1
    col2.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 3
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(3, 2, 1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.SameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(3, 2, 1)
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 1
    col2.Add 2
    col2.Add 3
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 1
    col2.Add 1
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 2
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 3
    col2.Add 2
    col2.Add 1
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(col2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)


    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(3, 2, 1)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.SameElementsAs(arr2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(3, 2, 1)
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(arr)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.SameElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").Should.Have.SameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs("""Hello world""")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.SameElementsAs(" ""Hello world"" ")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
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
    
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameElementsAs(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameElementsAs(col)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of("Hello world").ShouldNot.Have.SameElementsAs(d)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.SameElementsAs("Hello world")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.SameElementsAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.SameElementsAs("Hello world")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "SameElementsAsTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set SameElementsAsTests = testFluent
End Function

Private Function ProcedureTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests

    fluent.TestValue = testFluent.Of(testFluent).Should.Have.Procedure("Of", VbMethod)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative documentation tests
    
    fluent.TestValue = testFluent.Of(testFluent).ShouldNot.Have.Procedure("Of", VbMethod)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    
    'positive null documentation tests
    
    fluent.TestValue = testFluent.Of("Hello World").Should.Have.Procedure("Of", VbMethod)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
    fluent.TestValue = testFluent.Of("Hello World").ShouldNot.Have.Procedure("Of", VbMethod)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.Procedure("Of", VbMethod)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.Procedure("Of", VbMethod)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ProcedureTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ProcedureTests = testFluent
End Function

Private Function ElementsTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(1)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(1, 2)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(2, 3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(1, 3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.Elements(1, 2, 3)
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(1)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(1, 2)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(2, 3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(1, 3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.Elements(1, 2, 3)
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    fluent.TestValue = testFluent.Of(Null).Should.Have.Elements(1, 2, 3)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    'negative null documentation tests
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.Elements(1, 2, 3)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    fluent.TestValue = testFluent.Of().Should.Have.Elements(1, 2, 3)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.Elements(1, 2, 3)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ElementsTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ElementsTests = testFluent
End Function

Private Function ElementsInDataStructureTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(2))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(3))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(2, 3))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 3))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2, 3))
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        
    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(2))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(3))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(2, 3))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 3))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2, 3))
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    fluent.TestValue = testFluent.Of(Null).Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2, 3))
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2, 3))
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.ElementsInDataStructure(VBA.[_HiddenModule].Array(1, 2, 3))
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ElementsInDataStructureTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ElementsInDataStructureTests = testFluent
End Function

Private Function DepthCountOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim b As Boolean
    Dim arr() As Variant
    Dim d1 As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim d3 As Scripting.Dictionary
    Dim val As Variant
    Dim tfBitwiseFlag As cFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev
    
'positive documentation tests

    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(arr).Should.Have.DepthCountOf(0) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(0) = tfIter.Of(arr).Should.Have.DepthCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.DepthCountOf(1) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(1) = tfIter.Of(arr).Should.Have.DepthCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    fluent.TestValue = testFluent.Of(arr).Should.Have.DepthCountOf(2) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(2) = tfIter.Of(arr).Should.Have.DepthCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    fluent.TestValue = testFluent.Of(arr).Should.Have.DepthCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(3) = tfIter.Of(arr).Should.Have.DepthCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "C", 3
    d2.Add "B", d3
    d1.Add "A", d2
    
    fluent.TestValue = testFluent.Of(d1).Should.Have.DepthCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(d1).Should.Have.DepthCountOf(3) = tfIter.Of(d1).Should.Have.DepthCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "E", 5
    d3.Add "F", 6
    d2.Add "C", 3
    d2.Add "D", d3
    d1.Add "A", 1
    d1.Add "B", d2
    
    fluent.TestValue = testFluent.Of(d1).Should.Have.DepthCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(d1).Should.Have.DepthCountOf(3) = tfIter.Of(d1).Should.Have.DepthCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'negative documentation tests

    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.DepthCountOf(0) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(0) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.DepthCountOf(1) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(1) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.DepthCountOf(2) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(2) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.DepthCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(3) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "C", 3
    d2.Add "B", d3
    d1.Add "A", d2
    
    fluent.TestValue = testFluent.Of(d1).ShouldNot.Have.DepthCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(d1).ShouldNot.Have.DepthCountOf(3) = tfIter.Of(d1).ShouldNot.Have.DepthCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "E", 5
    d3.Add "F", 6
    d2.Add "C", 3
    d2.Add "D", d3
    d1.Add "A", 1
    d1.Add "B", d2
    
    fluent.TestValue = testFluent.Of(d1).ShouldNot.Have.DepthCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(d1).ShouldNot.Have.DepthCountOf(3) = tfIter.Of(d1).ShouldNot.Have.DepthCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'positive null documentation tests
    
    val = 0
    fluent.TestValue = testFluent.Of(Null).Should.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    val = 1
    fluent.TestValue = testFluent.Of(Null).Should.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 2
    fluent.TestValue = testFluent.Of(Null).Should.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 3
    fluent.TestValue = testFluent.Of(Null).Should.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'negative null documentation tests
    
    val = 0
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 1
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 2
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 3
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.DepthCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'empty documention tests
    
    val = 0
    fluent.TestValue = testFluent.Of().Should.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 1
    fluent.TestValue = testFluent.Of().Should.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 2
    fluent.TestValue = testFluent.Of().Should.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    val = 3
    fluent.TestValue = testFluent.Of().Should.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.DepthCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    
    Call validateRecurIterFluentOfs(testFluent, tfRecur, tfIter, "DepthCountOf")
    
    Call validateTests(fluent, testFluent)
    
'bitwise flags tests

    Set tfBitwiseFlag = MakeFluentOf
    tfBitwiseFlag.Meta.tests.Algorithm = flAlgorithm.flIterative + flAlgorithm.flRecursive
    
    arr = VBA.[_HiddenModule].Array(1)
    Debug.Assert tfBitwiseFlag.Of(arr).Should.Have.DepthCountOf(1)  'with implicit recur
    
    arr = VBA.[_HiddenModule].Array()
    Debug.Assert tfBitwiseFlag.Of(arr).ShouldNot.Have.DepthCountOf(1) 'with implicit recur
    
    Set testInfoDev = tfBitwiseFlag.Meta.tests.TestingFunctionsInfos
    
    With testInfoDev
        Debug.Assert .DepthCountOfRecur.Count > 0 And .DepthCountOfIter.Count > 0
        Debug.Assert .DepthCountOfRecur.Count = .DepthCountOfIter.Count
    End With
    
    Set tfBitwiseFlag = Nothing
    
    Debug.Print "DepthCountOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set DepthCountOfTests = testFluent
End Function

Private Function ErrorDescriptionOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
    
    fluent.TestValue = testFluent.Of("1 / 0").Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult) 'good
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        fluent.TestValue = testFluent.Of(Err).Should.Have.ErrorDescriptionOf("Division by zero")
    On Error GoTo 0
    
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
            
    'negative documentation tests
    
    fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        fluent.TestValue = testFluent.Of(Err).ShouldNot.Have.ErrorDescriptionOf("Division by zero")
    On Error GoTo 0
    
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'positive null documentation tests
    
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

    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'negative null documentation tests
    
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

    fluent.TestValue = testFluent.Of(VBA.[_HiddenModule].Array(1, 2, 3)).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New VBA.Collection
    fluent.TestValue = testFluent.Of(col).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    'empty documention tests
    
    fluent.TestValue = testFluent.Of().Should.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of().ShouldNot.Have.ErrorDescriptionOf("Application-defined or object-defined error")
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "ErrorDescriptionOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set ErrorDescriptionOfTests = testFluent
End Function

Private Function NestedCountOfTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim b As Boolean
    Dim arr() As Variant
    Dim val As Variant
    Dim tfBitwiseFlag As cFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev
    
'positive documentation tests

    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(0) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(0) = tfIter.Of(arr).Should.Have.NestedCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(1) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(1) = tfIter.Of(arr).Should.Have.NestedCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(2) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array())
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(0) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(0) = tfIter.Of(arr).Should.Have.NestedCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(1) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(1) = tfIter.Of(arr).Should.Have.NestedCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(2) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array()))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(0) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(0) = tfIter.Of(arr).Should.Have.NestedCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1)))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(1) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(1) = tfIter.Of(arr).Should.Have.NestedCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2)))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(2) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3)))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(2) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    fluent.TestValue = testFluent.Of(arr).Should.Have.NestedCountOf(3) 'with implicit recur
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'negative documentation tests

    arr = VBA.[_HiddenModule].Array()
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(0) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(0) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(1) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(1) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(2) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array())
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(0) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(0) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(1) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(1) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(2) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array()))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(0) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(0) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1)))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(1) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(1) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2)))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(2) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3)))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(2) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    fluent.TestValue = testFluent.Of(arr).ShouldNot.Have.NestedCountOf(3) 'with implicit recur
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'positive null documentation tests
    val = 0
    fluent.TestValue = testFluent.Of(Null).Should.Have.NestedCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.NestedCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
'negative null documentation tests
    
    fluent.TestValue = testFluent.Of(Null).ShouldNot.Have.NestedCountOf(val)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Have.NestedCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    
'empty documention tests
    
    val = 0
    fluent.TestValue = testFluent.Of().Should.Have.NestedCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.NestedCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).Should.Be.EqualTo(True) 'with explicit recur and iter
    Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    
    fluent.TestValue = testFluent.Of().ShouldNot.Have.NestedCountOf(val)
    Call EmptyAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.NestedCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    fluent.TestValue = testFluent.Of(b).ShouldNot.Be.EqualTo(True) 'with explicit recur and iter
    Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    
    Call validateRecurIterFluentOfs(testFluent, tfRecur, tfIter, "NestedCountOf")
    
    Call validateTests(fluent, testFluent)
    
'bitwise flags tests

    Set tfBitwiseFlag = MakeFluentOf
    tfBitwiseFlag.Meta.tests.Algorithm = flAlgorithm.flIterative + flAlgorithm.flRecursive
    
    arr = VBA.[_HiddenModule].Array(1)
    Debug.Assert tfBitwiseFlag.Of(arr).Should.Have.NestedCountOf(1)  'with implicit recur
    
    arr = VBA.[_HiddenModule].Array()
    Debug.Assert tfBitwiseFlag.Of(arr).ShouldNot.Have.NestedCountOf(1) 'with implicit recur
    
    Set testInfoDev = tfBitwiseFlag.Meta.tests.TestingFunctionsInfos
    
    With testInfoDev
        Debug.Assert .NestedCountOfRecur.Count > 0 And .NestedCountOfIter.Count > 0
        Debug.Assert .NestedCountOfRecur.Count = .NestedCountOfIter.Count
    End With
    
    Set tfBitwiseFlag = Nothing
    
    Debug.Print "NestedCountOfTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set NestedCountOfTests = testFluent
End Function



Private Function StubTests(ByVal fluent As IFluent, ByVal testFluent As IFluentOf, ByVal testFluentResult As IFluentOf) As IFluentOf
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    
    'positive documentation tests
            
    'negative documentation tests
    
    'positive null documentation tests
    
    'negative null documentation tests
    
    'empty documention tests
    
    Call validateTests(fluent, testFluent)
    
    Debug.Print "StubTests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set StubTests = testFluent
End Function

Private Function cleanStringTests(ByVal fluent As IFluent) As Long
    Dim testCount As Long
    
    fluent.Meta.tests.TestStrings.CleanTestValueStr = True
    
    fluent.TestValue = """abc"""
    
    Debug.Assert fluent.Should.Be.EqualTo("abc")
    
    fluent.Meta.tests.TestStrings.CleanTestValueStr = False
    fluent.Meta.tests.TestStrings.CleanTestInputStr = True
    
    fluent.TestValue = "abc"

    Debug.Assert fluent.Should.Be.EqualTo("""abc""")
    
    fluent.Meta.tests.TestStrings.CleanTestValueStr = True
    fluent.Meta.tests.TestStrings.CleanTestInputStr = True

    fluent.TestValue = """abc"""
    
    fluent.Meta.tests.TestStrings.CleanTestValueStr = False
    fluent.Meta.tests.TestStrings.CleanTestInputStr = False
    
    fluent.Meta.tests.TestStrings.CleanTestStrings = True

    fluent.TestValue = """abc"""

    Debug.Assert fluent.Should.Be.EqualTo("""abc""")
    
    'Add to clean strings tests
    
    fluent.Meta.tests.TestStrings.AddToCleanStringDict ("'")

    fluent.Meta.tests.TestStrings.CleanTestValueStr = True
    
    fluent.TestValue = "'abc def'"

    Debug.Assert fluent.Should.Be.EqualTo("abcdef")

    fluent.Meta.tests.TestStrings.CleanTestValueStr = False
    fluent.Meta.tests.TestStrings.CleanTestInputStr = True
    
    fluent.TestValue = "abcdef"
    
    Debug.Assert fluent.Should.Be.EqualTo("'abc def'")

    fluent.Meta.tests.TestStrings.CleanTestValueStr = False
    fluent.Meta.tests.TestStrings.CleanTestInputStr = True
    
    fluent.Meta.tests.TestStrings.AddToCleanStringDict " ", "_", True
    
    fluent.Meta.tests.TestStrings.CleanTestValueStr = True
    fluent.Meta.tests.TestStrings.CleanTestInputStr = False
    
    fluent.TestValue = "'abc def'"
    
    Debug.Assert fluent.Should.Be.EqualTo("abc_def")
    
    fluent.Meta.tests.TestStrings.CleanTestValueStr = False
    fluent.Meta.tests.TestStrings.CleanTestInputStr = True
    
    fluent.TestValue = "abc_def"
    
    Debug.Assert fluent.Should.Be.EqualTo("'abc def'")
    
    'Explicit clean strings using cUtilities
    
    fluent.Meta.tests.TestStrings.CleanTestStrings = False
    
    fluent.TestValue = """abc"""
    
    fluent.TestValue = fluent.Meta.tests.TestStrings.CleanString(fluent.TestValue)
    
    Debug.Assert fluent.Should.Be.EqualTo("abc")
    
    fluent.TestValue = """bcd"""
    
    fluent.TestValue = fluent.Meta.tests.TestStrings.CleanString(fluent.TestValue)
    
    Debug.Assert fluent.Should.Be.EqualTo(fluent.Meta.tests.TestStrings.CleanString("""bcd"""))
    
    Debug.Print "Clean string tests finished"
    
    testCount = fluent.Meta.tests.Count
    printTestCount (testCount)
    
    cleanStringTests = testCount
End Function

Private Function MiscTests(ByVal fluent As IFluent) As Long
    Dim testCount As Long
    Dim q As Object
    Dim elem As Variant
    Dim fluent2 As cFluent
    Dim col As Collection
    
    Set mEvents.setFluentEventDuplicate = fluent

    'test to ensure fluent object's default TestValue value is equal to empty
    Debug.Assert VBA.Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
    
    'test to ensure that a duplicate test event is not raised since skipDupCheck
    'is set to true
    
    With fluent.Meta.tests
        .SkipDupCheck = True
            Debug.Assert VBA.Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
        .SkipDupCheck = False
    End With
    
    'test to ensure that a duplicate test event is raised since skipDupCheck
    'is set to false
    
    Debug.Assert VBA.Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
    
    'test to ensure fluent object's TestValue property can return a value
    fluent.TestValue = fluent.TestValue
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    
    'test to ensure fluent object's TestValue property can return an object
    Set fluent.TestValue = New VBA.Collection
    Set fluent.TestValue = fluent.TestValue
    Debug.Assert fluent.Should.Be.Something
    
    'test to ensure that addDataStructure is working with non-default datastructure
    
    Set q = VBA.Interaction.CreateObject("system.collections.Queue")
    
    q.Enqueue ("Hello")
    
    fluent.Meta.tests.AddDataStructure q
    
    fluent.TestValue = "Hello"
    
    Debug.Assert fluent.Should.Be.InDataStructure(q)
    
    'test to ensure that StrTestValue and StrTestInput are working with non-default datastructure
    
    With fluent.Meta
        Debug.Assert .tests(.tests.Count).strTestValue = "`Hello`"
        Debug.Assert .tests(.tests.Count).StrTestInput = "Queue(`Hello`)"
    End With
    
    'Procedure bitwise flag tests
    
    Set fluent2 = MakeFluent
    
    Set fluent.TestValue = fluent2
    
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet)
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbGet)
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbSet)
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet + VbGet)
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet + VbSet)
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbGet + VbSet)
    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet + VbGet + VbSet)
    
    '//below tests will all fail since fluent objects do not have a TestValue method
    
    If Not G_TB_SKIP Then
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbGet + VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbSet + VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbGet + VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbSet + VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbGet + VbSet + VbMethod)
        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbGet + VbSet + VbMethod)
    End If
    
    '//self referential tests
    
    '//All self referential flags should be false
    
    Set col = New Collection
    
    Set fluent.TestValue = col
    
    fluent.Should.Be.Something
    
    With fluent.Meta
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).TestingValueIsSelfReferential)
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).TestingInputIsSelfReferential)
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).HasSelfReferential)
    End With
    
    col.Add col
    
    '//is null because col is now a self-referential data structure and
    '//ContinueWithSelfReferentialIfPossible is false
    Debug.Assert VBA.Information.IsNull(fluent.Should.Be.Something)
    
    fluent.Meta.tests.ContinueWithSelfReferentialIfPossible = True
    
    '//col is self-referential data structure but ContinueWithSelfReferentialIfPossible
    '//is set to true so the test now passes.
    
    Debug.Assert fluent.Should.Be.Something
    
    With fluent.Meta
        Debug.Assert .tests(.tests.Count).strTestValue = "Null"
        Debug.Assert .tests(.tests.Count).TestingValueIsSelfReferential = True
        Debug.Assert .tests(.tests.Count).HasSelfReferential = True
    End With
    
    '//testingValueIsSelfReferential, testingInputIsSelfReferential, and hasSelfReferential should be true
    
    Debug.Assert fluent.Should.Have.SameTypeAs(col)
    
    With fluent.Meta
        Debug.Assert .tests(.tests.Count).strTestValue = "Null"
        Debug.Assert .tests(.tests.Count).StrTestInput = "Null"
        Debug.Assert .tests(.tests.Count).TestingValueIsSelfReferential = True
        Debug.Assert .tests(.tests.Count).TestingInputIsSelfReferential = True
        Debug.Assert .tests(.tests.Count).HasSelfReferential = True
    End With
    
    Debug.Print "Misc tests finished"
    
    testCount = fluent.Meta.tests.Count
    printTestCount (testCount)
    
    MiscTests = testCount
End Function

Public Function checkResetCounters(ByVal fluent As IFluent, ByVal testFluent As IFluentOf) As Boolean
    Dim b As Boolean
    
    testFluent.Meta.tests.ResetCounter
    fluent.Meta.tests.ResetCounter
    
    b = (testFluent.Meta.tests.Count = 0 And fluent.Meta.tests.Count = 0)
   
   checkResetCounters = b
End Function

Public Function getFluentCounts(ByVal fluent As IFluent) As Boolean
    Dim test As ITest
    Dim d As Scripting.Dictionary
    Dim temp As String
    Dim elem As Variant
    Dim fn As String
    
    temp = ""
    Set d = New Scripting.Dictionary
    
    For Each test In fluent.Meta.tests
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

Public Function getFluentOfCounts(ByVal fluentOf As IFluentOf) As Boolean
    Dim test As ITest
    Dim d As Scripting.Dictionary
    Dim temp As String
    Dim elem As Variant
    Dim fn As String
    
    temp = ""
    Set d = New Scripting.Dictionary
    
    For Each test In fluentOf.Meta.tests
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

Private Function validateNegativeCounters(ByVal testFluent As IFluentOf) As Boolean
    Dim d As Scripting.Dictionary
    Dim test As ITest
    Dim counter As Long
    Dim fn As String
    Dim elem As Variant
    Dim testDev As ITestDev
    
    Set d = New Scripting.Dictionary
    
    counter = 0
    
    For Each test In testFluent.Meta.tests
        Set testDev = test
        If testDev.negateValue Then
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

Private Sub CheckTestFuncInfos(ByVal testFluent As IFluentOf)
    Dim d As Scripting.Dictionary
    Dim testFuncInfo As ITestingFunctionsInfo
    Dim testFuncInfoClass As cTestingFunctionsInfo
    Dim testFuncInfo2 As ITestingFunctionsInfo
    Dim testFuncInfoClass2 As cTestingFunctionsInfo
    Dim counter As Long
    
    Set d = testFluent.Meta.tests.TestingFunctionsInfos.TestFuncInfoToDict
    counter = 0
    
    With testFluent.Meta.tests
        For Each testFuncInfo In .TestingFunctionsInfos
            Set testFuncInfo2 = d(testFuncInfo.Name)
            
            Debug.Assert testFuncInfo.Name = testFuncInfo2.Name
            Debug.Assert testFuncInfo.Count = testFuncInfo2.Count
            Debug.Assert testFuncInfo.Failed = testFuncInfo2.Failed
            Debug.Assert testFuncInfo.Name = testFuncInfo2.Name
            Debug.Assert testFuncInfo.Passed = testFuncInfo2.Passed
            Debug.Assert testFuncInfo.Unexpected = testFuncInfo2.Unexpected
            
            counter = counter + 1
        Next testFuncInfo
        
        For Each testFuncInfoClass In .TestingFunctionsInfos
            Set testFuncInfoClass2 = d(testFuncInfoClass.Name)
            
            Debug.Assert testFuncInfoClass.Name = testFuncInfoClass2.Name
            Debug.Assert testFuncInfoClass.Count = testFuncInfoClass2.Count
            Debug.Assert testFuncInfoClass.Failed = testFuncInfoClass2.Failed
            Debug.Assert testFuncInfoClass.Name = testFuncInfoClass2.Name
            Debug.Assert testFuncInfoClass.Passed = testFuncInfoClass2.Passed
            Debug.Assert testFuncInfoClass.Unexpected = testFuncInfoClass2.Unexpected
            
            counter = counter + 1
        Next testFuncInfoClass
    End With
    
    'Debug.Print "CheckTestFuncInfos counter is: " & counter & vbNewLine
End Sub
