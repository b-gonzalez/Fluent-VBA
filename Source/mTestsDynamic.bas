Attribute VB_Name = "mTestsDynamic"
Option Explicit

Private Const G_TB_SKIP As Boolean = False

Private mCounter As Long
Private mTestCounter As Long
Private mMiscPosTests As Long
Private mMiscNegTests As Long
Private tfRecur As Variant
Private tfIter As Variant
Private mEvents As zEvents
Private mRecurIterFuncNamesDict As Scripting.Dictionary

Sub runMainTestsRefactor()
    Dim f As IFluent
    Dim fo As IFluentOf
    Dim f2 As IFluent
    Dim ff As IFluentFunction
    Dim tf As IFluentOf
    Dim temp As Variant
    Dim fiArr As Variant
    Dim i As Long
    Dim counter As Long
    
    Set fo = New cFluentOf
    Set f2 = New cFluent
    Set ff = New cFluentFunction
    
    fiArr = Array(fo, f2, ff)
'    fiArr = Array(ff)

    For i = LBound(fiArr) To UBound(fiArr)
        mCounter = 0
        
        'Creating new instances of f and tf in the loop is
        'necessary. Otherwise the counts will be incorrect
        'in getAndInitEventRefactor in the second element
        'of the fiArr array which will make the test fail.
        
        Set f = New cFluent
        Set tf = New cFluentOf
        'This is needed for validateRecurIterFluentOfs
        Set mRecurIterFuncNamesDict = New Scripting.Dictionary
        
        'Creating new instances of tfRecur and tfIter in the
        'loop is necessary. Otherwise, the checks in
        'validateRecurIterFluentOfsRefactor will fail since
        'tfRecur and tfIter will maintain their counts from
        'the previous element in the array while the new
        'element in fiArr will not. And so, their counts will
        'be different and one of the tests will fail.
        
        Set tfRecur = MakeFluentOf
        Set tfIter = MakeFluentOf
    
        tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
        tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
        
        Set fiArr(i) = InitFluentInput(fiArr(i))
        Set mEvents = getAndInitEventRefactor(f, fiArr(i), tf)
        
        Set fiArr(i) = AlphabeticTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = AlphanumericTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = BetweenTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = ContainTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = DepthCountOfTestsRefactor(f, fiArr(i), tf)
        
        Set fiArr(i) = ElementsInDataStructureTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = ElementsTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = EndWithTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = EqualityDocumentationTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = EqualToTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = EvaluateToTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = ExactSameElementsAsTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = GreaterThanTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = GreaterThanOrEqualToTestsRfactor(f, fiArr(i), tf)
        Set fiArr(i) = IdenticalToTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = InDataStructureTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = InDataStructuresTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = LengthBetweenTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = LengthOfTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = LessThanTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = LessThanOrEqualToTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = MaxLengthOfTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = MinLengthOfTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = NestedCountOfTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = NumericTestsReactor(f, fiArr(i), tf)
        Set fiArr(i) = OneOfTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = ProcedureTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = StartWithTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = SomethingTestsReactor(f, fiArr(i), tf)
        Set fiArr(i) = SameTypeAsTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = SameUniqueElementsAsTestsRefactor(f, fiArr(i), tf)
        Set fiArr(i) = SameElementsAsTestsRefactor(f, fiArr(i), tf)
    
        If Not G_TB_SKIP Then
            Set fiArr(i) = ErroneousTestsRefactor(f, fiArr(i), tf)
            Set fiArr(i) = ErrorDescriptionOfTestsRefactor(f, fiArr(i), tf)
            Set fiArr(i) = ErrorNumberOfTestsRefactor(f, fiArr(i), tf)
            counter = 0
        Else
            counter = 3
        End If
        
        Call runRecurIterTestsRefactor(fiArr(i))
        Call cleanStringTestsRefactor(fiArr(i))
        Call MiscTestsRefactor(fiArr(i))
        
        Set mEvents = Nothing
        Set tfRecur = Nothing
        Set tfIter = Nothing
        Set mRecurIterFuncNamesDict = Nothing
    Next i
    
    Debug.Print "All refactored tests finished!"
End Sub

Public Sub TrueAssertAndRaiseEventsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf)
    Dim td As ITestDev
    Dim inputIter As String
    Dim inputRecur As String
    Dim valueIter As String
    Dim valueRecur As String
    Dim testFluent As IFluent
    Dim testFluentOf As IFluentOf
    Dim testFluentFunction As IFluentFunction

    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    If TypeOf fluentInput Is cFluent Or TypeOf fluentInput Is IFluent Then
        Set testFluent = fluentInput
        Debug.Assert testFluent.Meta.tests.Count = mCounter
        fluent.testValue = testFluent.Meta.tests(mCounter).result '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentOf Or TypeOf fluentInput Is IFluentOf Then
        Set testFluentOf = fluentInput
        Debug.Assert testFluentOf.Meta.tests.Count = mCounter
        fluent.testValue = testFluentOf.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentFunction Or TypeOf fluentInput Is IFluentFunction Then
        Set testFluentFunction = fluentInput
        Debug.Assert testFluentFunction.Meta.tests.Count = mCounter
        fluent.testValue = testFluentFunction.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    End If

    With fluent.Meta.tests
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    If fluentInput.Meta.tests.ToStrDev Then
        With fluentInput.Meta.tests
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

Public Sub FalseAssertAndRaiseEventsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf)
    Dim td As ITestDev
    Dim inputIter As String
    Dim inputRecur As String
    Dim valueIter As String
    Dim valueRecur As String
    Dim testFluent As IFluent
    Dim testFluentOf As IFluentOf
    Dim testFluentFunction As IFluentFunction
    
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1

    If TypeOf fluentInput Is cFluent Or TypeOf fluentInput Is IFluent Then
        Set testFluent = fluentInput
        Debug.Assert testFluent.Meta.tests.Count = mCounter
        fluent.testValue = testFluent.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentOf Or TypeOf fluentInput Is IFluentOf Then
        Set testFluentOf = fluentInput
        Debug.Assert testFluentOf.Meta.tests.Count = mCounter
        fluent.testValue = testFluentOf.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentFunction Or TypeOf fluentInput Is IFluentFunction Then
        Set testFluentFunction = fluentInput
        Debug.Assert testFluentFunction.Meta.tests.Count = mCounter
        fluent.testValue = testFluentFunction.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    End If
    
    With fluent.Meta.tests
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With

    If fluentInput.Meta.tests.ToStrDev Then
        With fluentInput.Meta.tests
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

Private Sub NullAssertAndRaiseEventsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf)
    Dim testFluent As IFluent
    Dim testFluentOf As IFluentOf
    Dim testFluentFunction As IFluentFunction
    
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    If TypeOf fluentInput Is cFluent Or TypeOf fluentInput Is IFluent Then
        Set testFluent = fluentInput
        Debug.Assert testFluent.Meta.tests.Count = mCounter
        fluent.testValue = testFluent.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentOf Or TypeOf fluentInput Is IFluentOf Then
        Set testFluentOf = fluentInput
        Debug.Assert testFluentOf.Meta.tests.Count = mCounter
        fluent.testValue = testFluentOf.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentFunction Or TypeOf fluentInput Is IFluentFunction Then
        Set testFluentFunction = fluentInput
        Debug.Assert testFluentFunction.Meta.tests.Count = mCounter
        fluent.testValue = testFluentFunction.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    End If
    
'    Debug.Assert fluentInput.Meta.tests.Count = mCounter

    With fluent
        Debug.Assert testFluentResult.Of(.testValue).Should.Be.EqualTo(Null)
        Debug.Assert testFluentResult.Of(.testValue).ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.testValue).ShouldNot.Be.EqualTo(False)
    End With
End Sub

Private Sub EmptyAssertAndRaiseEventsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf)
    Dim testFluent As IFluent
    Dim testFluentOf As IFluentOf
    Dim testFluentFunction As IFluentFunction

    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1

'    Debug.Assert testFluent.Meta.Tests(testFluent.Meta.Tests.Count).TestValueSet = False
    If TypeOf fluentInput Is cFluent Or TypeOf fluentInput Is IFluent Then
        Set testFluent = fluentInput
        Debug.Assert testFluent.Meta.tests.Count = mCounter
        fluent.testValue = testFluent.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentOf Or TypeOf fluentInput Is IFluentOf Then
        Set testFluentOf = fluentInput
        Debug.Assert testFluentOf.Meta.tests.Count = mCounter
        fluent.testValue = testFluentOf.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    ElseIf TypeOf fluentInput Is cFluentFunction Or TypeOf fluentInput Is IFluentFunction Then
        Set testFluentFunction = fluentInput
        Debug.Assert testFluentFunction.Meta.tests.Count = mCounter
        fluent.testValue = testFluentFunction.Meta.tests(mCounter).result  '//comment out until all tests are refactored unless testing.
    End If
    
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
'    Debug.Assert fluentInput.Meta.tests.Count = mCounter

    Debug.Assert fluentInput.Meta.tests(mCounter).TestValueSet = False

    With fluent
        Debug.Assert VBA.Information.IsEmpty(fluent.testValue)
    End With
End Sub

Private Function EqualityDocumentationTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim test As ITest
    Dim i As Long
    Dim resultBool As Boolean
    Dim fluentBool As Boolean
    Dim valueBool As Boolean
    Dim inputBool As Boolean
    Dim counter As Long
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    counter = 0

    With fluent.Meta.tests
    
        testingValue = True
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = True
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = False
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = False
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = True
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = True
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = False
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = False
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = -1
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = -1
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = 0
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = 0
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = -1
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = -1
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = 0
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = 0
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        '//Approximate equality tests
    
        fluentInput.Meta.tests.ApproximateEqual = True
        testingValue = "TRUE"
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "TRUE"
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "FALSE"
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "FALSE"
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = "TRUE"
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "TRUE"
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "FALSE"
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "FALSE"
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "true"
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "true"
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "false"
        testingInput = True
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "false"
        testingInput = False
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = "true"
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "true"
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "false"
        testingInput = True
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
        testingValue = "false"
        testingInput = False
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        fluentInput.Meta.tests.ApproximateEqual = False
        
        '//Null and Empty tests
        
        testingValue = Null
        testingInput = Null
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = Null
        testingInput = Null
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = Empty
        testingInput = Empty
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = Empty
        testingInput = Empty
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = ""
        testingInput = Empty
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = ""
        testingInput = Empty
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = 0
        testingInput = Empty
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = 0
        testingInput = Empty
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = False
        testingInput = Empty
        shouldMatch = True
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        testingValue = False
        testingInput = Empty
        shouldMatch = False
        functionName = "EqualTo"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue, testingInput)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    End With
    
    For Each test In fluent.Meta.tests
        Debug.Assert test.result
    Next test
    
    For i = 1 To fluent.Meta.tests.Count
        Debug.Assert fluent.Meta.tests(i).result
    Next i
    
    i = 1
    
    With fluentInput.Meta
        For Each test In .tests
            If Not VBA.Information.IsNull(test.result) And Not VBA.Information.IsNull(.tests(i).result) Then
                resultBool = test.result = .tests(i).result
                fluentBool = test.FluentPath = .tests(i).FluentPath
                valueBool = test.testingValue = .tests(i).testingValue
                inputBool = test.testingInput = .tests(i).testingInput
                
                Debug.Assert resultBool And fluentBool And valueBool And inputBool
                
                i = i + 1
            End If
        Next test
    End With
    
    Debug.Print "Equality tests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set EqualityDocumentationTestsRefactor = fluentInput
End Function

Private Function GreaterThanTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
        
    testingValue = 10
    testingInput = 9
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 11
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 9
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null tests
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documentation tests
    
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)
    
    Debug.Print "GreaterThanTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set GreaterThanTestsRefactor = fluentInput
End Function

Private Function EqualToTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    testingValue = """abc"""
    testingInput = """abc"""
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    ' //Approximate equality tests
    fluentInput.Meta.tests.ApproximateEqual = True
    testingValue = "10"
    testingInput = 10
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "True"
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    testingValue = 5.0000001
    testingInput = 5
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = CStr(Excel.Evaluate("1 / 0"))
    testingInput = "Error 2007"
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc"""
    testingInput = """abc"""
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 10
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = CStr(Excel.Evaluate("1 / 0"))
    testingInput = "Error 2007"
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    ' //Approximate equality tests
    fluentInput.Meta.tests.ApproximateEqual = True
    
    testingValue = "10"
    testingInput = 10
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "True"
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    testingValue = 5.0000001
    testingInput = 5
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty tests

    testingInput = "Hello world"
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "EqualToTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set EqualToTestsRefactor = fluentInput

End Function

Private Function GreaterThanOrEqualToTestsRfactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput = 9
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 9.1
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 11
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 11.1
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput = 9
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 9.1
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 11
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 11.1
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = 10
    shouldMatch = True
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 10
    shouldMatch = False
    functionName = "GreaterThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "GreaterThanOrEqualToTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set GreaterThanOrEqualToTestsRfactor = fluentInput
End Function

Private Function LessThanTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    testingValue = 10
    testingInput = 9
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 9
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'empty documentation tests
    testingInput = 10
    shouldMatch = True
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 10
    shouldMatch = False
    functionName = "LessThan"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "LessThanTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set LessThanTestsRefactor = fluentInput
End Function

Private Function LessThanOrEqualToTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput = 9
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 10.1
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 11.1
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 9.1
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput = 9
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 10.1
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 11.1
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10.1
    testingInput = 9.1
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = 10
    shouldMatch = True
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 10
    shouldMatch = False
    functionName = "LessThanOrEqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "LessThanOrEqualToTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set LessThanOrEqualToTestsRefactor = fluentInput
End Function


Private Function ContainTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    testingValue = 10
    testingInput = 1
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 0
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 2
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "10"
    testingInput = "1"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "0"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "10"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "2"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = "Hello"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = "world"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "ru"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "1"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "0"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "10"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "2"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello world"""
    testingInput = "Hello"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello world"""
    testingInput = "world"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "ru"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "als"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative documentation tests
    testingValue = 10
    testingInput = 1
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 0
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 2
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "1"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "0"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "10"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "2"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "Hello world"
    testingInput = "Hello"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = "world"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "1"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "0"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "10"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """10"""
    testingInput = "2"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello world"""
    testingInput = "Hello"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello world"""
    testingInput = "world"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "ru"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "als"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abcde"""
    testingInput = "abc"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world!"""
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """Hello world!"""
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """Hello world!"""
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """Hello world!"""
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world!"" "
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = " ""Hello world!"" "
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = " ""Hello world!"" "
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = " ""Hello world!"" "
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world! """
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """ Hello world! """
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """ Hello world! """
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """ Hello world! """
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world!"""
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """Hello world!"""
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """Hello world!"""
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """Hello world!"""
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world!"" "
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = " ""Hello world!"" "
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = " ""Hello world!"" "
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = " ""Hello world!"" "
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world! """
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """ Hello world! """
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """ Hello world! """
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """ Hello world! """
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'empty tests
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "Contain"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ContainTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ContainTestsRefactor = fluentInput
End Function

Private Function StartWithTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    testingValue = 10
    testingInput = 1
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 10
    testingInput = 2
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "1"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "10"
    testingInput = "2"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello World"
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "1 ""0"" "
    testingInput = "1"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "1 ""0"" "
    testingInput = "2"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello ""World"" "
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "True"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "T"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "False"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "F"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative documentation tests
    testingValue = 10
    testingInput = 1
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 2
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "1"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "2"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "Hello World"
    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "1 ""0"" "
    testingInput = "1"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "1 ""0"" "
    testingInput = "2"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello ""World"" "
    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "True"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "T"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "False"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "F"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documentation tests
    testingInput = "Hello"
    shouldMatch = True
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello"
    shouldMatch = False
    functionName = "StartWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "StartWithTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set StartWithTestsRefactor = fluentInput
End Function

Private Function EndWithTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput = 0
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 2
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "0"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "2"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello World"
    testingInput = "World"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = " ""1"" 0"
    testingInput = "0"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""1"" 0"
    testingInput = "2"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""Hello"" World"
    testingInput = "World"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "True"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "e"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "False"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "e"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative documentation tests
    
    testingValue = 10
    testingInput = 0
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 2
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "0"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = "2"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello World"
    testingInput = "World"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = " ""1"" 0"
    testingInput = "0"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""1"" 0"
    testingInput = "2"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""Hello"" World"
    testingInput = "World"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "True"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = "e"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "False"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = "e"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    testingValue = Null
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = "Hello world!"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = """Hello"""
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = " ""Hello"" "
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """ Hello """
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = Null
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    testingValue = Null
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = "Hello world!"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = """Hello"""
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = " ""Hello"" "
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = """ Hello """
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'empty documention tests
    testingInput = "Hello"
    shouldMatch = True
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello"
    shouldMatch = False
    functionName = "EndWith"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "EndWithTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set EndWithTestsRefactor = fluentInput
End Function

Private Function LengthOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    testingInput = 3
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = 4
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = Len("10")
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    testingInput = Len("abc")
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = Len("True")
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 1
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative documentation tests
    
    testingValue = 10
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    testingInput = 3
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = 4
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = Len("10")
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    testingInput = Len("abc")
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = Len("True")
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 1
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = 2
    shouldMatch = True
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 2
    shouldMatch = False
    functionName = "LengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "LengthOfTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set LengthOfTestsRefactor = fluentInput
End Function

Private Function MaxLengthOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput = 3
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 1
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = "10"
    testingInput = 3
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 1
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = Len("True")
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = Len("False")
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput = 3
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 1
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 3
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 1
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = Len("True")
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = Len("False")
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = 2
    shouldMatch = True
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 2
    shouldMatch = False
    functionName = "MaxLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "MaxLengthOfTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set MaxLengthOfTestsRefactor = fluentInput
End Function

Private Function MinLengthOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput = 3
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 1
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 3
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 1
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = Len("True")
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = Len("False")
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput = 3
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 1
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 3
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "10"
    testingInput = 1
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = Len("True")
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = Len("False")
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add "Hello world!"
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """Hello world!"""
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add " ""Hello world!"" "
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    col.Add """ Hello world! """
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array()
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = VBA.Interaction.CreateObject("Scripting.Dictionary")
    Set testingValue = d
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    Set testingValue = d
    testingInput = 2.34
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests

    testingInput = 2
    shouldMatch = True
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = 2
    shouldMatch = False
    functionName = "MinLengthOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "MinLengthOfTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set MinLengthOfTestsRefactor = fluentInput
End Function


Private Function BetweenTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput1 = 10
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9.99
    testingInput2 = 10.01
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9
    testingInput2 = 11
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9.1
    testingInput2 = 11.1
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 11
    testingInput2 = 9
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 11.1
    testingInput2 = 9.1
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput1 = 10
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9.99
    testingInput2 = 10.01
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9
    testingInput2 = 11
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9.1
    testingInput2 = 11.1
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 11
    testingInput2 = 9
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 11.1
    testingInput2 = 9.1
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = "Hello World!"
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello World!"""
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""Hello World!"" "
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ Hello World! """
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative null documentation tests
    
    testingValue = "Hello World!"
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello World!"""
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""Hello World!"" "
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ Hello World! """
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "Between"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "BetweenTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set BetweenTestsRefactor = fluentInput
End Function



Private Function LengthBetweenTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput1 = 1
    testingInput2 = 3
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 0
    testingInput2 = 2
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 2
    testingInput2 = 2
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 3
    testingInput2 = 1
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 2
    testingInput2 = 0
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput1 = 1
    testingInput2 = 3
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 0
    testingInput2 = 2
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 2
    testingInput2 = 2
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 3
    testingInput2 = 1
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 2
    testingInput2 = 0
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = True
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput1 = 1
    testingInput2 = 10
    shouldMatch = False
    functionName = "LengthBetween"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, lowerVal:=testingInput1, higherVal:=testingInput2)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "LengthBetweenTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set LengthBetweenTestsRefactor = fluentInput
End Function

Private Function OneOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim testingInput3 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = 10
    testingInput1 = 9
    testingInput2 = 10
    testingInput3 = 11
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 10
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9
    testingInput2 = 11
    testingInput3 = 13
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    ' //Object and data structure tests
    
    Set col = New VBA.Collection
    Set d = New Scripting.Dictionary
    Set testingValue = col
    Set testingInput1 = col
    Set testingInput2 = d
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    testingValue = 10
    Set testingInput1 = col
    Set testingInput2 = d
    testingInput3 = 10
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = 10
    testingInput = 10
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9
    testingInput2 = 10
    testingInput3 = 11
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput1 = 9
    testingInput2 = 11
    testingInput3 = 13
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    testingInput = 11
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 10
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    ' //Object and data structure tests
    
    Set col = New VBA.Collection
    Set d = New Scripting.Dictionary
    Set testingValue = col
    Set testingInput1 = col
    Set testingInput2 = d
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    testingValue = 10
    Set testingInput1 = col
    Set testingInput2 = d
    testingInput3 = 10
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'OneOf does not have positive or negative null documentation tests
        
    'empty documention tests
    
    testingInput1 = "Hello world"
    testingInput2 = 5
    testingInput3 = True
    shouldMatch = True
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput1 = "Hello world"
    testingInput2 = 5
    testingInput3 = True
    shouldMatch = False
    functionName = "OneOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "OneOfTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set OneOfTestsRefactor = fluentInput
End Function

Private Function SomethingTestsReactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = Nothing
    Set testingValue = col
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = Nothing
    Set testingValue = col
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
                        
    testingValue = "Hello World!"
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello World!"""
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""Hello World!"" "
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ Hello World! """
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 123
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 1.23
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = "Hello World!"
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """Hello World!"""
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""Hello World!"" "
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ Hello World! """
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 123
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 1.23
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    shouldMatch = False
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "SomethingTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set SomethingTestsReactor = fluentInput
End Function

Private Function EvaluateToTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim arr() As Variant
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    testingValue = True
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "true"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "false"
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "TRUE"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "FALSE"
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = -1
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = -1
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 0
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 0
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "-1"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "-1"
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "0"
    testingInput = False
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "0"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 5 + 5
    testingInput = 10
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "5 + 5"
    testingInput = 10
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "5 + 5 = 10"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "5 + 5 > 9"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array()
    testingValue = VBA.Information.TypeName(arr) = "Variant()"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.Information.IsArray(arr)
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = VBA.Information.TypeName(col) = "Collection"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = TypeOf col Is Collection
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = VBA.Information.TypeName(d) = "Dictionary"
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = TypeOf d Is Scripting.Dictionary
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    '//Testing errors is possible if they're put in strings
    testingValue = "1 / 0"
    testingInput = CVErr(xlErrDiv0)
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'negative documentation tests
                
    testingValue = True
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "true"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "false"
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "TRUE"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "FALSE"
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = -1
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = -1
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 0
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 0
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "-1"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "-1"
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "0"
    testingInput = False
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "0"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 5 + 5
    testingInput = 10
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "5 + 5"
    testingInput = 10
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "5 + 5 = 10"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "5 + 5 > 9"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array()
    testingValue = VBA.Information.TypeName(arr) = "Variant()"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.Information.IsArray(arr)
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = VBA.Information.TypeName(col) = "Collection"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = TypeOf col Is Collection
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = VBA.Information.TypeName(d) = "Dictionary"
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = TypeOf d Is Scripting.Dictionary
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    '//Testing errors is possible if they're put in strings
    testingValue = "1 / 0"
    testingInput = CVErr(xlErrDiv0)
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = True
    shouldMatch = True
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = True
    shouldMatch = False
    functionName = "EvaluateTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "EvaluateToTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set EvaluateToTestsRefactor = fluentInput
End Function

Private Function AlphabeticTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
        
    testingValue = "abc"
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc!@#"
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123"
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "!@#"
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True
    
    testingValue = "abc def"
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " abc "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " abc!@# "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " 123 "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " !@# "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc"""
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc!@#"""
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123"""
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """!@#"""
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc"" "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc!@#"" "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123"" "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""!@#"" "
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc """
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc!@# """
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123 """
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ !@# """
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False

    'negative documentation tests
    
    testingValue = "abc"
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc!@#"
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123"
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "!@#"
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True
    
    testingValue = "abc def"
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " abc "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " abc!@# "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " 123 "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " !@# "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = """abc"""
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc!@#"""
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123"""
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """!@#"""
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc"" "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc!@#"" "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123"" "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""!@#"" "
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc """
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc!@# """
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123 """
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ !@# """
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    shouldMatch = True
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    shouldMatch = False
    functionName = "Alphabetic"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "AlphabeticTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set AlphabeticTestsRefactor = fluentInput
End Function

Private Function NumericTestsReactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    
    testingValue = 123
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123"
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123!@#"
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "!@#"
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True
    
    testingValue = "123 456"
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123"""
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123!@#"""
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc"""
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """!@#"""
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123"" "
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123!@#"" "
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc"" "
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""!@#"" "
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123 """
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123!@# """
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc """
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ !@# """
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False
    
    'negative documentation tests
    
    testingValue = 123
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123"
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123!@#"
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = "abc"
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "!@#"
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True
    
    testingValue = "123 456"
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123"""
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123!@#"""
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = """abc"""
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """!@#"""
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123"" "
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123!@#"" "
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = " ""abc"" "
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""!@#"" "
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123 """
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123!@# """
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    testingValue = """ abc """
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ !@# """
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    shouldMatch = True
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    shouldMatch = False
    functionName = "Numeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "NumericTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set NumericTestsReactor = fluentInput
End Function

Private Function AlphanumericTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    
    testingValue = "abc123"
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123"
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "!@#"
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True
    
    testingValue = "abc 123"
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc123"""
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc"""
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123"""
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """!@#"""
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc123"" "
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc"" "
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123"" "
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""!@#"" "
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc123 """
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc """
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123 """
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ !@# """
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False
    
    'negative documentation tests
    
    testingValue = "abc123"
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "abc"
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "123"
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "!@#"
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True
    
    testingValue = "abc 123"
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc123"""
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """abc"""
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """123"""
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """!@#"""
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc123"" "
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""abc"" "
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""123"" "
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = " ""!@#"" "
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc123 """
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ abc """
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ 123 """
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = """ !@# """
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False

    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    shouldMatch = True
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    shouldMatch = False
    functionName = "Alphanumeric"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "AlphanumericTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set AlphanumericTestsRefactor = fluentInput
End Function

Private Function SameTypeAsTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim arr() As Variant
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    testingValue = CBool(True)
    testingInput = CBool(True)
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("Hello World!")
    testingInput = CStr("Goodbye World!")
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("""Hello World!""")
    testingInput = CStr("""Goodbye World!""")
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = CStr("""Hello World!""")
    testingInput = CStr("Goodbye World!")
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("Hello World!")
    testingInput = CStr("""Goodbye World!""")
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CLng(12345)
    testingInput = CLng(54321)
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CSng(123.45)
    testingInput = CSng(543.21)
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CDbl(123.45)
    testingInput = CDbl(543.21)
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CDate(#12/31/1999#)
    testingInput = CDate(#12/31/2000#)
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = Nothing
    Set testingInput = Nothing
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    Set testingInput = d
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    Set testingValue = d
    Set testingInput = d
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CLng(123)
    testingInput = CStr("Hello world")
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CLng(123)
    testingInput = CDbl(123.456)
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = CLng(123)
    Set testingInput = col
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing

    'negative documentation tests
    testingValue = CBool(True)
    testingInput = CBool(True)
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("Hello World!")
    testingInput = CStr("Goodbye World!")
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("""Hello World!""")
    testingInput = CStr("""Goodbye World!""")
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("""Hello World!""")
    testingInput = CStr("Goodbye World!")
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CStr("Hello World!")
    testingInput = CStr("""Goodbye World!""")
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CLng(12345)
    testingInput = CLng(54321)
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CSng(123.45)
    testingInput = CSng(543.21)
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CDbl(123.45)
    testingInput = CDbl(543.21)
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CDate(#12/31/1999#)
    testingInput = CDate(#12/31/2000#)
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = Nothing
    Set testingInput = Nothing
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    Set testingInput = col
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    Set testingValue = col
    Set testingInput = col
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    Set testingInput = d
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    Set testingValue = d
    Set testingInput = d
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = CLng(123)
    testingInput = CStr("Hello world")
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = CLng(123)
    testingInput = CDbl(123.456)
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = CLng(123)
    Set testingInput = col
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    'SameTypeAs does not have positive or negative null documentation tests
    
    'empty documention tests
    
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameTypeAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "SameTypeAsTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set SameTypeAsTestsRefactor = fluentInput
End Function

Private Function IdenticalToTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col).Should.Be
'        testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col).Should.Be
'        testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1
    
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col2
    Set testingInput = col3
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2
    Set testingValue = col
    Set testingInput = col3
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col2
    Set testingInput = col3
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col2
    Set testingInput = col
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col2).Should.Be
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = Nothing
    Set col2 = Nothing
    Set col3 = Nothing
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
'    With testFluent.Of(arr).Should.Be
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
'    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).Should.Be
'        testingValue = arr
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col).ShouldNot.Be
'        testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col).ShouldNot.Be
'        testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1
    
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col2
    Set testingInput = col3
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2
    Set testingValue = col
    Set testingInput = col3
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col2
    Set testingInput = col3
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col2
    Set testingInput = col
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col2).ShouldNot.Be
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = Nothing
    Set col2 = Nothing
    Set col3 = Nothing
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
'    With testFluent.Of(arr).ShouldNot.Be
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    
'    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).ShouldNot.Be
'        testingValue = arr
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    'positive null documentation tests
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'    testingValue = "Hello world"
'    testingInput1 = VBA.[_HiddenModule].Array(1
'    testingInput2 = 2
'    testingInput3 =  3)
'    shouldMatch = True
'    functionName = "IdenticalTo"
'    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
'    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = Null
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'    testingValue = "Hello world"
'    testingInput1 = VBA.[_HiddenModule].Array(1
'    testingInput2 = 2
'    testingInput3 =  3)
'    shouldMatch = False
'    functionName = "IdenticalTo"
'    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
'    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = Null
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello world"
    shouldMatch = False
    functionName = "IdenticalTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "IdenticalToTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set IdenticalToTestsRefactor = fluentInput
End Function

Private Function ExactSameElementsAsTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col).Should.Have
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1

'    With testFluent.Of(col).Should.Have
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1

'    With testFluent.Of(col2).Should.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2

'    With testFluent.Of(col3).Should.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col2
    Set testingInput = col
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1

'    With testFluent.Of(col2).Should.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = Nothing
    Set col2 = Nothing
    Set col3 = Nothing
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

'    With testFluent.Of(arr).Should.Have
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

'    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).Should.Have
'        testingValue = arr
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1)
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(2)
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1

'    With testFluent.Of(VBA.[_HiddenModule].Array(1)).Should.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1

'    With testFluent.Of(col).Should.Have
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2

'    With testFluent.Of(col2).Should.Have
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2

'    With testFluent.Of(col2).Should.Have
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col2
    Set testingInput = col
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    

    Set col = New VBA.Collection
    col.Add 1
    
'    With testFluent.Of(VBA.[_HiddenModule].Array(2)).Should.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = Nothing

    
    'negative documentation tests
    
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1
    
'    With testFluent.Of(col).ShouldNot.Have
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1
    col3.Add 1

'    With testFluent.Of(col).ShouldNot.Have
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2
    col3.Add 1

'    With testFluent.Of(col2).ShouldNot.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 2

'    With testFluent.Of(col3).ShouldNot.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col2
    Set testingInput = col
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 1
    col3.Add 1

'    With testFluent.Of(col2).ShouldNot.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = col3
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = Nothing
    Set col2 = Nothing
    Set col3 = Nothing
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

'    With testFluent.Of(arr).ShouldNot.Have
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 3, 4)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = arr

'    With testFluent.Of(VBA.[_HiddenModule].Array(2, 3, 4)).ShouldNot.Have
'        testingValue = arr
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = arr2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1)
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(2)
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1

'    With testFluent.Of(VBA.[_HiddenModule].Array(1)).ShouldNot.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 2
    col2.Add 1

'    With testFluent.Of(col).ShouldNot.Have
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col2
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2

'    With testFluent.Of(col2).ShouldNot.Have
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col3 = New VBA.Collection
    col.Add 1
    col2.Add 2

'    With testFluent.Of(col2).ShouldNot.Have
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col2
    Set testingInput = col
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
'    With testFluent.Of(VBA.[_HiddenModule].Array(2)).ShouldNot.Have
'        Set testingValue = col
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'        testingValue = VBA.[_HiddenModule].Array(1)
'        None
'        functionName = "None"
'        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
'        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
'    End With
    Set col = Nothing
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello world"
    shouldMatch = False
    functionName = "ExactSameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ExactSameElementsAsTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ExactSameElementsAsTestsRefactor = fluentInput
End Function

Private Function SameUniqueElementsAsTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    col.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 3
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    col.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 0)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 2
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 0)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    col2.Add 2
    col2.Add 1
    col.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 1)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 1
    col2.Add 3
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 2
    col2.Add 1
    col.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 0)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(2, 1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    


    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 2
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1, 0)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(2, 1, 2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameUniqueElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "SameUniqueElementsAsTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set SameUniqueElementsAsTestsRefactor = fluentInput
End Function

Private Function SameElementsAsTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim col2 As VBA.Collection
    Dim col3 As VBA.Collection
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 1
    col2.Add 2
    col2.Add 3
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    testingValue = arr
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 1
    col2.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 3
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(3, 2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(3, 2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    Set testingValue = col
    Set testingInput = col
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 1
    col2.Add 2
    col2.Add 3
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)


    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)


    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(1, 2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(1, 2)
    testingValue = arr
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    Set testingValue = col
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 1
    col2.Add 1
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col2.Add 2
    col2.Add 2
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    
    Set col = New VBA.Collection
    Set col2 = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col2.Add 3
    col2.Add 2
    col2.Add 1
    Set testingValue = col
    Set testingInput = col2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(1)
    arr2 = VBA.[_HiddenModule].Array(2)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)


    arr = VBA.[_HiddenModule].Array(1, 2)
    arr2 = VBA.[_HiddenModule].Array(2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    arr2 = VBA.[_HiddenModule].Array(3, 2, 1)
    testingValue = arr
    testingInput = arr2
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    arr = VBA.[_HiddenModule].Array(2)
    Set testingValue = col
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    arr = VBA.[_HiddenModule].Array(2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    arr = VBA.[_HiddenModule].Array(3, 2, 1)
    Set testingValue = col
    testingInput = arr
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    'positive null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = "Hello world"
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    testingValue = "Hello world"
    Set testingInput = col
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    testingValue = "Hello world"
    Set testingInput = d
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = "Hello world"
    shouldMatch = True
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Hello world"
    shouldMatch = False
    functionName = "SameElementsAs"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "SameElementsAsTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set SameElementsAsTestsRefactor = fluentInput
End Function

Private Function ProcedureTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    
    If TypeOf fluentInput Is IFluentOf Or TypeOf fluentInput Is cFluentOf Then
        Set testingValue = fluentInput
        testingInput1 = "Of"
        testingInput2 = VbMethod
        shouldMatch = True
        functionName = "Procedure"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
        Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
        'negative documentation tests
        
        Set testingValue = fluentInput
        testingInput1 = "Of"
        testingInput2 = VbMethod
        shouldMatch = False
        functionName = "Procedure"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
        Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    End If
    'positive null documentation tests
    
    testingValue = "Hello World"
    testingInput1 = "Of"
    testingInput2 = VbMethod
    shouldMatch = True
    functionName = "Procedure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = "Hello World"
    testingInput1 = "Of"
    testingInput2 = VbMethod
    shouldMatch = False
    functionName = "Procedure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput1 = "Of"
    testingInput2 = VbMethod
    shouldMatch = True
    functionName = "Procedure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput1 = "Of"
    testingInput2 = VbMethod
    shouldMatch = False
    functionName = "Procedure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ProcedureTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ProcedureTestsRefactor = fluentInput
End Function

Private Function ElementsTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim testingInput3 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    Set testingValue = col
    testingInput = 1
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = 2
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = 3
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 2
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 2
    testingInput2 = 3
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 3
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 2
    testingInput3 = 3
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    Set testingValue = col
    testingInput = 1
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = 2
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = 3
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 2
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 2
    testingInput2 = 3
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 3
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput1 = 1
    testingInput2 = 2
    testingInput3 = 3
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    testingValue = Null
    testingInput1 = 1
    testingInput2 = 2
    testingInput3 = 3
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    'negative null documentation tests
    testingValue = Null
    testingInput1 = 1
    testingInput2 = 2
    testingInput3 = 3
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    testingInput1 = 1
    testingInput2 = 2
    testingInput3 = 3
    shouldMatch = True
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput1 = 1
    testingInput2 = 2
    testingInput3 = 3
    shouldMatch = False
    functionName = "Elements"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ElementsTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ElementsTestsRefactor = fluentInput
End Function

Private Function ElementsInDataStructureTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    'Dim testingInput1 As Variant
    'Dim testingInput2 As Variant
    'Dim testingInput3 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(2)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(3)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1, 2)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(2, 3)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1, 3)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

        
    'negative documentation tests
    Set col = New VBA.Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(2)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(3)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1, 2)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(2, 3)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1, 3)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set testingValue = col
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    testingValue = Null
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    testingValue = Null
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "ElementsInDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ElementsInDataStructureTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ElementsInDataStructureTestsRefactor = fluentInput
End Function

Private Function InDataStructureTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim col2 As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim arr() As Variant
    Dim strArr(1, 1) As Variant
    Dim b As Boolean
    Dim al As Object
    Dim val As Variant
    Dim tfBitwiseFlag As Variant
    Dim testInfoDev As ITestingFunctionsInfoDev
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
'positive documentation tests
    
    arr = VBA.[_HiddenModule].Array()
    testingValue = 10
    testingInput = arr
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    testingValue = b
    testingInput = False
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(9, 10, 11)
    testingValue = 10
    testingInput = arr
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    ReDim arr(1, 1)
    arr(0, 0) = 9
    arr(0, 1) = 10
    arr(1, 0) = 11
    arr(1, 1) = 12
    testingValue = 10
    testingInput = arr
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    ReDim arr(1, 1, 1)
    arr(0, 0, 0) = 6
    arr(0, 0, 1) = 7
    arr(0, 1, 0) = 8
    arr(0, 1, 1) = 9
    arr(1, 0, 0) = 10
    arr(1, 0, 1) = 11
    arr(1, 1, 0) = 12
    arr(1, 1, 1) = 13
    testingValue = 10
    testingInput = arr
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    testingValue = 10
    testingInput = arr
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(arr) = tfIter.Of(10).Should.Be.InDataStructure(arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add 10
    col.Add 11
    testingValue = 10
    Set testingInput = col
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(col) = tfIter.Of(10).Should.Be.InDataStructure(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = col
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(col) = tfIter.Of(10).Should.Be.InDataStructure(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testingValue = 10
    Set testingInput = d
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(d) = tfIter.Of(10).Should.Be.InDataStructure(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = d
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(d) = tfIter.Of(10).Should.Be.InDataStructure(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 9, 1
    d.Add 10, 2
    d.Add 11, 3
    testingValue = 10
    testingInput = d.Keys
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(d.Keys) = tfIter.Of(10).Should.Be.InDataStructure(d.Keys)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructure(al) = tfIter.Of(10).Should.Be.InDataStructure(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
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
    
    testingValue = 1
    Set testingInput = d
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(1).Should.Be.InDataStructure(d) = tfIter.Of(1).Should.Be.InDataStructure(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'negative documentation tests
    
    arr = VBA.[_HiddenModule].Array()
    testingValue = 10
    testingInput = arr
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = False
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    arr = VBA.[_HiddenModule].Array(9, 10, 11)
    testingValue = 10
    testingInput = arr
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    ReDim arr(1, 1)
    arr(0, 0) = 9
    arr(0, 1) = 10
    arr(1, 0) = 11
    arr(1, 1) = 12
    testingValue = 10
    testingInput = arr
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    ReDim arr(1, 1, 1)
    arr(0, 0, 0) = 6
    arr(0, 0, 1) = 7
    arr(0, 1, 0) = 8
    arr(0, 1, 1) = 9
    arr(1, 0, 0) = 10
    arr(1, 0, 1) = 11
    arr(1, 1, 0) = 12
    arr(1, 1, 1) = 13
    testingValue = 10
    testingInput = arr
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    testingValue = 10
    testingInput = arr
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(arr) = tfIter.Of(10).ShouldNot.Be.InDataStructure(arr)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add 10
    col.Add 11
    testingValue = 10
    Set testingInput = col
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(col) = tfIter.Of(10).ShouldNot.Be.InDataStructure(col)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = col
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(col) = tfIter.Of(10).ShouldNot.Be.InDataStructure(col)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set col = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testingValue = 10
    Set testingInput = d
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(d) = tfIter.Of(10).ShouldNot.Be.InDataStructure(d)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set d = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = d
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(d) = tfIter.Of(10).ShouldNot.Be.InDataStructure(d)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set d = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 9, 1
    d.Add 10, 2
    d.Add 11, 3
    testingValue = 10
    testingInput = d.Keys
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(d.Keys) = tfIter.Of(10).ShouldNot.Be.InDataStructure(d.Keys)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set d = Nothing
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructure(al) = tfIter.Of(10).ShouldNot.Be.InDataStructure(al)
    mMiscNegTests = mMiscNegTests + 1
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

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
    
    testingValue = 1
    Set testingInput = d
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(1).ShouldNot.Be.InDataStructure(d) = tfIter.Of(1).ShouldNot.Be.InDataStructure(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'positive null documentation tests

    val = "Hello World"
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """ Hello World """
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 10
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 123.45
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = True
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = Null
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'negative null documentation tests
    
    val = "Hello World"
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = """ Hello World """
    testingValue = 10
    testingInput = """ Hello World """
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = 10
    testingValue = 10
    testingInput = 10
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 123.45
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = True
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = Null
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'empty documention tests
    
    val = "Hello World"
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """Hello world"""
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = " ""Hello world"" "
    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """ Hello world """
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = val
    shouldMatch = False
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructure(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructure(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    
    Call validateRecurIterFluentOfsRefactor(fluentInput, tfRecur, tfIter, "InDataStructure")

    Call validateTestsRefactor(fluent, fluentInput)
    
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
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set InDataStructureTestsRefactor = fluentInput
End Function

Private Function InDataStructuresTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim b As Boolean
    Dim al As Object
    Dim val As Variant
    Dim tfBitwiseFlag As IFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim testingInput3 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

'positive documentation tests
           
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    testingValue = 10
    testingInput = arr2
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(arr2) = tfIter.Of(10).Should.Be.InDataStructures(arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    ReDim arr(1, 1)
    arr(0, 0) = 12
    arr(0, 1) = 13
    arr(1, 0) = 14
    arr(1, 1) = 15
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    testingValue = 12
    testingInput1 = arr
    testingInput2 = arr2
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    b = tfRecur.Of(12).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(12).Should.Be.InDataStructures(arr, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

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
    testingValue = 9
    testingInput1 = arr
    testingInput2 = arr2
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(9).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(9).Should.Be.InDataStructures(arr, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    testingValue = 10
    testingInput1 = arr
    testingInput2 = arr2
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(arr, arr2) = tfIter.Of(10).Should.Be.InDataStructures(arr, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    testingValue = 13
    Set testingInput = col
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(13).Should.Be.InDataStructures(col) = tfIter.Of(13).Should.Be.InDataStructures(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    testingValue = 16
    testingInput1 = arr
    Set testingInput2 = col
    testingInput3 = arr2
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(16).Should.Be.InDataStructures(arr, col, arr2) = tfIter.Of(16).Should.Be.InDataStructures(arr, col, arr2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
                    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = col
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(col) = tfIter.Of(10).Should.Be.InDataStructures(col)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing
    
    arr = VBA.[_HiddenModule].Array(12, 13, 14)
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 14
    Set testingInput1 = col
    testingInput2 = arr
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(14).Should.Be.InDataStructures(col, arr) = tfIter.Of(14).Should.Be.InDataStructures(col, arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set col = Nothing

    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testingValue = 10
    Set testingInput = d
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(d) = tfIter.Of(10).Should.Be.InDataStructures(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testingValue = 2
    testingInput1 = d.Items
    testingInput2 = d.Keys
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(2).Should.Be.InDataStructures(d.Items, d.Keys) = tfIter.Of(2).Should.Be.InDataStructures(d.Items, d.Keys)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = Nothing
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = d
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(d) = tfIter.Of(10).Should.Be.InDataStructures(d)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(6, VBA.[_HiddenModule].Array(7, VBA.[_HiddenModule].Array(8)))
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 8
    Set testingInput1 = al
    testingInput2 = arr
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(8).Should.Be.InDataStructures(al, arr) = tfIter.Of(8).Should.Be.InDataStructures(al, arr)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'negative documentation tests
    
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    testingValue = 10
    testingInput = arr2
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(arr2) = tfIter.Of(10).ShouldNot.Be.InDataStructures(arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    ReDim arr(1, 1)
    arr(0, 0) = 12
    arr(0, 1) = 13
    arr(1, 0) = 14
    arr(1, 1) = 15
    arr2 = VBA.[_HiddenModule].Array(9, 10, 11)
    testingValue = 12
    testingInput1 = arr
    testingInput2 = arr2
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(12).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(12).ShouldNot.Be.InDataStructures(arr, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

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
    testingValue = 9
    testingInput1 = arr
    testingInput2 = arr2
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(9).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(9).ShouldNot.Be.InDataStructures(arr, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    testingValue = 10
    testingInput1 = arr
    testingInput2 = arr2
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(arr, arr2) = tfIter.Of(10).ShouldNot.Be.InDataStructures(arr, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    testingValue = 13
    Set testingInput = col
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(13).ShouldNot.Be.InDataStructures(col) = tfIter.Of(13).ShouldNot.Be.InDataStructures(col)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    
    Set col = New VBA.Collection
    col.Add 12
    col.Add 13
    col.Add 14
    arr = VBA.[_HiddenModule].Array(9, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11)))
    arr2 = VBA.[_HiddenModule].Array(15, 16, 17)
    testingValue = 16
    testingInput1 = arr
    Set testingInput2 = col
    testingInput3 = arr2
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2, testingInput3:=testingInput3)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2) = tfIter.Of(16).ShouldNot.Be.InDataStructures(arr, col, arr2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = col
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(col) = tfIter.Of(10).ShouldNot.Be.InDataStructures(col)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set col = Nothing

    arr = VBA.[_HiddenModule].Array(12, 13, 14)
    Set col = New VBA.Collection
    col.Add 9
    col.Add VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 14
    Set testingInput1 = col
    testingInput2 = arr
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(14).ShouldNot.Be.InDataStructures(col, arr) = tfIter.Of(14).ShouldNot.Be.InDataStructures(col, arr)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set col = Nothing

    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testingValue = 10
    Set testingInput = d
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(d) = tfIter.Of(10).ShouldNot.Be.InDataStructures(d)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    Set d = Nothing ''

    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testingValue = 2
    testingInput1 = d.Items
    testingInput2 = d.Keys
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(2).ShouldNot.Be.InDataStructures(d.Items, d.Keys) = tfIter.Of(2).ShouldNot.Be.InDataStructures(d.Items, d.Keys)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set d = Nothing
    Set d = New Scripting.Dictionary
    d.Add 1, 9
    d.Add 2, VBA.[_HiddenModule].Array(10, VBA.[_HiddenModule].Array(11))
    testingValue = 10
    Set testingInput = d
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(d) = tfIter.Of(10).ShouldNot.Be.InDataStructures(d)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''
    Set d = Nothing

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).ShouldNot.Be.InDataStructures(al) = tfIter.Of(10).ShouldNot.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    arr = VBA.[_HiddenModule].Array(6, VBA.[_HiddenModule].Array(7, VBA.[_HiddenModule].Array(8)))
    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 8
    Set testingInput1 = al
    testingInput2 = arr
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(8).ShouldNot.Be.InDataStructures(al, arr) = tfIter.Of(8).ShouldNot.Be.InDataStructures(al, arr)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) ''

    Set al = VBA.Interaction.CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testingValue = 10
    Set testingInput = al
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(10).Should.Be.InDataStructures(al) = tfIter.Of(10).Should.Be.InDataStructures(al)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

'positive null documentation tests
    
    val = "Hello World"
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """Hello World"""
    testingValue = 10
    testingInput = """Hello World"""
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = " ""Hello World"" "
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """ Hello World """
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = 10
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = 123.45
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = True
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = Null
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = "Hello World"
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """Hello World"""
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = " ""Hello World "" "
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """ Hello World """
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 10
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 123.45
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = True
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = Null
    testingValue = 10
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).Should.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'negative null documentation tests
    
    val = "Hello World"
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """Hello World"""
    testingValue = 10
    testingInput = """Hello World"""
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = " ""Hello World"" "
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """ Hello World """
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = 10
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = 123.45
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = True
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = Null
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = "Hello World"
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """Hello World"""
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = " ""Hello World "" "
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = """ Hello World """
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 10
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 123.45
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = True
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = Null
    testingValue = 10
    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'empty documention tests
    
    val = "Hello World"
    testingInput = val
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = val
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """Hello world"""
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """Hello world"""
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = " ""Hello world"" "
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = " ""Hello world"" "
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """ Hello world """
    shouldMatch = True
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = """ Hello world """
    shouldMatch = False
    functionName = "InDataStructures"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Be.InDataStructures(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Be.InDataStructures(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive

    Call validateRecurIterFluentOfsRefactor(fluentInput, tfRecur, tfIter, "InDataStructures")
'
    Call validateTestsRefactor(fluent, fluentInput)
    
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
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set InDataStructuresTestsRefactor = fluentInput
End Function

Private Function DepthCountOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim b As Boolean
    Dim arr() As Variant
    Dim d1 As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim d3 As Scripting.Dictionary
    Dim val As Variant
    Dim tfBitwiseFlag As IFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev
    Dim testingValue As Variant
    Dim testingInput As Variant
'    Dim testingInput1 As Variant
'    Dim testingInput2 As Variant
'    Dim testingInput3 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    
'positive documentation tests

    arr = VBA.[_HiddenModule].Array()
    testingValue = arr
    testingInput = 0
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(0) = tfIter.Of(arr).Should.Have.DepthCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = 1
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(1) = tfIter.Of(arr).Should.Have.DepthCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    testingValue = arr
    testingInput = 2
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(2) = tfIter.Of(arr).Should.Have.DepthCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    testingValue = arr
    testingInput = 3
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.DepthCountOf(3) = tfIter.Of(arr).Should.Have.DepthCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "C", 3
    d2.Add "B", d3
    d1.Add "A", d2
    
    Set testingValue = d1
    testingInput = 3
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(d1).Should.Have.DepthCountOf(3) = tfIter.Of(d1).Should.Have.DepthCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "E", 5
    d3.Add "F", 6
    d2.Add "C", 3
    d2.Add "D", d3
    d1.Add "A", 1
    d1.Add "B", d2
    
    Set testingValue = d1
    testingInput = 3
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(d1).Should.Have.DepthCountOf(3) = tfIter.Of(d1).Should.Have.DepthCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
'negative documentation tests

    arr = VBA.[_HiddenModule].Array()
    testingValue = arr
    testingInput = 0
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(0) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = 1
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(1) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    testingValue = arr
    testingInput = 2
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(2) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    testingValue = arr
    testingInput = 3
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.DepthCountOf(3) = tfIter.Of(arr).ShouldNot.Have.DepthCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "C", 3
    d2.Add "B", d3
    d1.Add "A", d2
    
    Set testingValue = d1
    testingInput = 3
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(d1).ShouldNot.Have.DepthCountOf(3) = tfIter.Of(d1).ShouldNot.Have.DepthCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d1 = New Scripting.Dictionary
    Set d2 = New Scripting.Dictionary
    Set d3 = New Scripting.Dictionary
    
    d3.Add "E", 5
    d3.Add "F", 6
    d2.Add "C", 3
    d2.Add "D", d3
    d1.Add "A", 1
    d1.Add "B", d2
    
    Set testingValue = d1
    testingInput = 3
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(d1).ShouldNot.Have.DepthCountOf(3) = tfIter.Of(d1).ShouldNot.Have.DepthCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'positive null documentation tests
    
    val = 0
    testingValue = Null
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    val = 1
    testingValue = Null
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 2
    testingValue = Null
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 3
    testingValue = Null
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'negative null documentation tests
    
    val = 0
    testingValue = Null
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 1
    testingValue = Null
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 2
    testingValue = Null
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 3
    testingValue = Null
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'empty documention tests
    
    val = 0
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 1
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 2
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    val = 3
    testingInput = val
    shouldMatch = True
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingInput = val
    shouldMatch = False
    functionName = "DepthCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.DepthCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.DepthCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    
    Call validateRecurIterFluentOfsRefactor(fluentInput, tfRecur, tfIter, "DepthCountOf")
    
    Call validateTestsRefactor(fluent, fluentInput)
    
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
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set DepthCountOfTestsRefactor = fluentInput
End Function

Private Function NestedCountOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim b As Boolean
    Dim arr() As Variant
    Dim val As Variant
    Dim tfBitwiseFlag As IFluentOf
    Dim testInfoDev As ITestingFunctionsInfoDev
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
'positive documentation tests

    arr = VBA.[_HiddenModule].Array()
    testingValue = arr
    testingInput = 0
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(0) = tfIter.Of(arr).Should.Have.NestedCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = 1
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(1) = tfIter.Of(arr).Should.Have.NestedCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    testingValue = arr
    testingInput = 2
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = 3
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array())
    testingValue = arr
    testingInput = 0
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(0) = tfIter.Of(arr).Should.Have.NestedCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1))
    testingValue = arr
    testingInput = 1
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(1) = tfIter.Of(arr).Should.Have.NestedCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2))
    testingValue = arr
    testingInput = 2
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3))
    testingValue = arr
    testingInput = 3
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array()))
    testingValue = arr
    testingInput = 0
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(0) = tfIter.Of(arr).Should.Have.NestedCountOf(0)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1)))
    testingValue = arr
    testingInput = 1
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(1) = tfIter.Of(arr).Should.Have.NestedCountOf(1)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2)))
    testingValue = arr
    testingInput = 2
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3)))
    testingValue = arr
    testingInput = 3
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    testingValue = arr
    testingInput = 2
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(2) = tfIter.Of(arr).Should.Have.NestedCountOf(2)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    testingValue = arr
    testingInput = 3
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).Should.Have.NestedCountOf(3) = tfIter.Of(arr).Should.Have.NestedCountOf(3)
    mMiscPosTests = mMiscPosTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
        
'negative documentation tests

    arr = VBA.[_HiddenModule].Array()
    testingValue = arr
    testingInput = 0
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(0) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1)
    testingValue = arr
    testingInput = 1
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(1) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2)
    testingValue = arr
    testingInput = 2
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, 2, 3)
    testingValue = arr
    testingInput = 3
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array())
    testingValue = arr
    testingInput = 0
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(0) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1))
    testingValue = arr
    testingInput = 1
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(1) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2))
    testingValue = arr
    testingInput = 2
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3))
    testingValue = arr
    testingInput = 3
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array()))
    testingValue = arr
    testingInput = 0
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(0) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(0)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1)))
    testingValue = arr
    testingInput = 1
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(1) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(1)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2)))
    testingValue = arr
    testingInput = 2
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(VBA.[_HiddenModule].Array(1, 2, 3)))
    testingValue = arr
    testingInput = 3
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2))
    testingValue = arr
    testingInput = 2
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(2) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(2)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    arr = VBA.[_HiddenModule].Array(1, VBA.[_HiddenModule].Array(2, VBA.[_HiddenModule].Array(3)))
    testingValue = arr
    testingInput = 3
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = tfRecur.Of(arr).ShouldNot.Have.NestedCountOf(3) = tfIter.Of(arr).ShouldNot.Have.NestedCountOf(3)
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'positive null documentation tests
    val = 0
    testingValue = Null
    testingInput = val
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(Null).Should.Have.NestedCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(Null).Should.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
'negative null documentation tests
    
    testingValue = Null
    testingInput = val
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsNull(tfRecur.Of(10).ShouldNot.Have.NestedCountOf(val)) And _
    VBA.Information.IsNull(tfIter.Of(10).ShouldNot.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    
'empty documention tests
    
    val = 0
    testingInput = val
    shouldMatch = True
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().Should.Have.NestedCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().Should.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    
    testingInput = val
    shouldMatch = False
    functionName = "NestedCountOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    b = _
    VBA.Information.IsEmpty(tfRecur.Of().ShouldNot.Have.NestedCountOf(val)) And _
    VBA.Information.IsEmpty(tfIter.Of().ShouldNot.Have.NestedCountOf(val))
    mMiscNegTests = mMiscNegTests + 1 'incrementing misc counter to account for second test in b
    testingValue = b
    testingInput = True
    shouldMatch = False
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    
    Debug.Assert tfIter.Meta.tests.Algorithm = flAlgorithm.flIterative
    Debug.Assert tfRecur.Meta.tests.Algorithm = flAlgorithm.flRecursive
    
    Call validateRecurIterFluentOfsRefactor(fluentInput, tfRecur, tfIter, "NestedCountOf")
    
    Call validateTestsRefactor(fluent, fluentInput)
    
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
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set NestedCountOfTestsRefactor = fluentInput
End Function

Private Function ErroneousTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    
    testingValue = "1 / 0"
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        Set testingValue = Err
        shouldMatch = True
        functionName = "Erroneous"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    On Error GoTo 0
    
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) 'good
                
    'negative documentation tests
    
    testingValue = "1 / 0"
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        Set testingValue = Err
        shouldMatch = False
        functionName = "Erroneous"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    On Error GoTo 0
    
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = 123
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 1.23
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = 123
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = 1.23
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = True
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = False
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = Null
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set col = New VBA.Collection
    Set testingValue = col
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Set d = New Scripting.Dictionary
    Set testingValue = d
    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    shouldMatch = True
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    shouldMatch = False
    functionName = "Erroneous"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ErroneousTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ErroneousTestsRefactor = fluentInput
End Function

Private Function ErrorDescriptionOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    'positive documentation tests
    
    testingValue = "1 / 0"
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult) 'good
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        Set testingValue = Err
        testingInput = "Division by zero"
        shouldMatch = True
        functionName = "ErrorDescriptionOf"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    On Error GoTo 0
    
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
            
    'negative documentation tests
    
    testingValue = "1 / 0"
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        Set testingValue = Err
        testingInput = "Division by zero"
        shouldMatch = False
        functionName = "ErrorDescriptionOf"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    On Error GoTo 0
    
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = 123
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 1.23
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = True
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = False
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = Null
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = 123
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 1.23
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = True
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = False
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = Null
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = "Application-defined or object-defined error"
    shouldMatch = True
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "Application-defined or object-defined error"
    shouldMatch = False
    functionName = "ErrorDescriptionOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)
    
    Debug.Print "ErrorDescriptionOfTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ErrorDescriptionOfTestsRefactor = fluentInput
End Function

Private Function ErrorNumberOfTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As Variant
    Dim col As VBA.Collection
    Dim d As Scripting.Dictionary
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String

    'positive documentation tests
    
    testingValue = "1 / 0"
    testingInput = 2007
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        Set testingValue = Err
        testingInput = 11
        shouldMatch = True
        functionName = "ErrorNumberOf"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    On Error GoTo 0
    
    Call TrueAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative documentation tests
    
    testingValue = "1 / 0"
    testingInput = 2007
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    On Error Resume Next
        Debug.Print 1 / 0
    
        Set testingValue = Err
        testingInput = 11
        shouldMatch = False
        functionName = "ErrorNumberOf"
        Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    On Error GoTo 0
    
    Call FalseAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'positive null documentation tests
    
    testingValue = 123
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 1.23
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = True
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = False
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = Null
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'negative null documentation tests
    
    testingValue = 123
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = 1.23
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = True
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = False
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = Null
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingValue = VBA.[_HiddenModule].Array(1, 2, 3)
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set col = New VBA.Collection
    Set testingValue = col
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    Set d = New Scripting.Dictionary
    Set testingValue = d
    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Call NullAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    'empty documention tests
    
    testingInput = "2007"
    shouldMatch = True
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)

    testingInput = "2007"
    shouldMatch = False
    functionName = "ErrorNumberOf"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Call EmptyAssertAndRaiseEventsRefactor(fluent, fluentInput, testFluentResult)
    
    Call validateTestsRefactor(fluent, fluentInput)

    Debug.Print "ErrorNumberOfTests finished"
    printTestCountRefactor (mTestCounter)
    mTestCounter = 0

    Set ErrorNumberOfTestsRefactor = fluentInput
End Function

Private Function cleanStringTestsRefactor(ByVal fluentInput As Variant) As Long
    Dim testCount As Long
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = True
    
    testingValue = """abc"""
    
    testingInput = "abc"
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    
    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = False
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = True
    
    testingValue = "abc"

    testingInput = """abc"""
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    
    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = True
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = True

    testingValue = """abc"""
    
    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = False
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = False
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = True

    testingValue = """abc"""

    testingInput = """abc"""
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    
    'Add to clean strings tests
    
    fluentInput.Meta.tests.TestStrings.AddToCleanStringDict ("'")

    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = True
    
    testingValue = "'abc def'"

    testingInput = "abcdef"
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)

    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = False
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = True
    
    testingValue = "abcdef"
    
    testingInput = "'abc def'"
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)

    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = False
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = True
    
    fluentInput.Meta.tests.TestStrings.AddToCleanStringDict " ", "_", True
    
    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = True
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = False
    
    testingValue = "'abc def'"
    
    testingInput = "abc_def"
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    
    fluentInput.Meta.tests.TestStrings.CleanTestValueStr = False
    fluentInput.Meta.tests.TestStrings.CleanTestInputStr = True
    
    testingValue = "abc_def"
    
    testingInput = "'abc def'"
    
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    
    'Explicit clean strings using cUtilities
    
    fluentInput.Meta.tests.TestStrings.CleanTestStrings = False
    
    testingValue = """abc"""
    
    testingInput = fluentInput.Meta.tests.TestStrings.CleanString(testingValue)
    
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
    testingValue = """bcd"""
    
    testingInput = fluentInput.Meta.tests.TestStrings.CleanString(testingValue)
    
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Print "Clean string tests finished"
'
'    testCount = fluentInput.Meta.tests.Count
'    printTestCount (testCount)
'
'    cleanStringTestsRefactor = testCount
End Function

Private Function MiscTestsRefactor(ByVal fluentInput As Variant) As Long
    Dim testCount As Long
    Dim q As Object
    Dim elem As Variant
    Dim fluent2 As IFluent
    Dim col As Collection
    Dim testingValue As Variant
    Dim testingInput As Variant
    Dim testingInput1 As Variant
    Dim testingInput2 As Variant
    Dim shouldMatch As Boolean
    Dim functionName As String
    
'    Set mEvents.setFluentEventDuplicate = fluent

    'test to ensure fluent object's default TestValue value is equal to empty
'    Debug.Assert VBA.Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
    
    testingInput = Empty
    shouldMatch = True
    functionName = "EqualTo"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingInput:=testingInput)
    Debug.Assert VBA.Information.IsEmpty(fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result)
    
    'test to ensure that a duplicate test event is not raised since skipDupCheck
    'is set to true
    
'    With fluent.Meta.tests
'        .SkipDupCheck = True
'            Debug.Assert VBA.Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
'        .SkipDupCheck = False
'    End With
    
    'test to ensure that a duplicate test event is raised since skipDupCheck
    'is set to false
    
'    Debug.Assert VBA.Information.IsEmpty(fluent.Should.Be.EqualTo(Empty))
    
    'test to ensure fluent object's TestValue property can return a value
'    testingValue = fluentInput.Meta.tests(fluentInput.Meta.tests.Count).testingValue
'    testingInput = Empty
'    shouldMatch = True
'    functionName = "EqualTo"
'    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
'    Debug.Assert VBA.Information.IsEmpty(fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result)
    
    'test to ensure fluent object's TestValue property can return an object
'    Set fluent.testValue = New VBA.Collection
'    Set fluent.testValue = fluent.testValue
'    Debug.Assert fluent.Should.Be.Something

    Set testingValue = New VBA.Collection
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Debug.Assert VBA.Information.IsObject(fluentInput.Meta.tests(fluentInput.Meta.tests.Count).testingValue)
    
    'test to ensure that addDataStructure is working with non-default datastructure
    
    Set q = VBA.Interaction.CreateObject("system.collections.Queue")
    
    q.Enqueue ("Hello")
    
    fluentInput.Meta.tests.AddDataStructure q
    
'    fluent.testValue = "Hello"
'
'    Debug.Assert fluent.Should.Be.InDataStructure(q)

    testingValue = "Hello"
    Set testingInput = q
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
    'test to ensure that StrTestValue and StrTestInput are working with non-default datastructure
    
    With fluentInput.Meta
        Debug.Assert .tests(.tests.Count).strTestValue = "`Hello`"
        Debug.Assert .tests(.tests.Count).StrTestInput = "Queue(`Hello`)"
    End With
    
    'Procedure bitwise flag tests
    
    Set fluent2 = MakeFluent
    
'    Set fluent.testValue = fluent2
    Set testingValue = fluent2
    testingInput1 = "TestValue"
    functionName = "Procedure"
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet)
    testingInput2 = VbLet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbGet)
    testingInput2 = VbGet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbSet)
    testingInput2 = VbSet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet + VbGet)
    testingInput2 = VbLet + VbGet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet + VbSet)
    testingInput2 = VbLet + VbSet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbGet + VbSet)
    testingInput2 = VbGet + VbSet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
'    Debug.Assert fluent.Should.Have.Procedure("TestValue", VbLet + VbGet + VbSet)
    testingInput2 = VbLet + VbGet + VbSet
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput1:=testingInput1, testingInput2:=testingInput2)
    Debug.Assert fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result
    
    '//below tests will all fail since fluent objects do not have a TestValue method
    
'    If Not G_TB_SKIP Then
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbGet + VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbSet + VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbGet + VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbSet + VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbGet + VbSet + VbMethod)
'        Debug.Assert fluent.ShouldNot.Have.Procedure("TestValue", VbLet + VbGet + VbSet + VbMethod)
'    End If
    
    '//self referential tests
    
    '//All self referential flags should be false
    
    Set col = New Collection
'    Debug.Assert fluent.Should.Be.Something
    
    Set testingValue = col
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Debug.Assert VBA.Information.IsObject(fluentInput.Meta.tests(fluentInput.Meta.tests.Count).testingValue)

    
    '//Checks that all properties self-referential properties should be Null since neither
    '//the testing value nor testing input is self referential
    
    With fluentInput.Meta
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).TestingValueIsSelfReferential)
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).TestingInputIsSelfReferential)
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).HasSelfReferential)
    End With
    
    Set col = New VBA.Collection
    
    col.Add 1
    col.Add col
    
'    set fluent.testValue = col
'    Debug.Assert VBA.Information.IsNull(fluent.Should.Be.Something)
    
    Set testingValue = col
    shouldMatch = True
    functionName = "Something"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue)
    Debug.Assert VBA.Information.IsNull(fluentInput.Meta.tests(fluentInput.Meta.tests.Count).testingValue)
    
    '//testingValueIsSelfReferential and hasSelfReferential should be true
    
    With fluentInput.Meta
        Debug.Assert .tests(.tests.Count).strTestValue = "Null"
        Debug.Assert .tests(.tests.Count).TestingValueIsSelfReferential = True
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).TestingInputIsSelfReferential)
        Debug.Assert .tests(.tests.Count).HasSelfReferential = True
    End With
    
    '//testingValueIsSelfReferential, testingInputIsSelfReferential, and hasSelfReferential should be true
    
'    fluent.testValue = 1
'
'    Debug.Assert VBA.Information.IsNull(fluent.Should.Be.InDataStructure(col))

    testingValue = 1
    Set testingInput = col
    shouldMatch = True
    functionName = "InDataStructure"
    Set fluentInput = fluentTester(fluentInput, functionName, shouldMatch, testingValue:=testingValue, testingInput:=testingInput)
    Debug.Assert VBA.Information.IsNull(fluentInput.Meta.tests(fluentInput.Meta.tests.Count).result)
    
    With fluentInput.Meta
        Debug.Assert Not VBA.Information.IsNull(.tests(.tests.Count).strTestValue)
        Debug.Assert .tests(.tests.Count).StrTestInput = "Null"
        Debug.Assert VBA.Information.IsNull(.tests(.tests.Count).TestingValueIsSelfReferential)
        Debug.Assert .tests(.tests.Count).TestingInputIsSelfReferential = True
        Debug.Assert .tests(.tests.Count).HasSelfReferential = True
    End With
    
'    Debug.Print "Misc tests finished"
'
'    testCount = fluent.Meta.tests.Count
'    printTestCount (testCount)
'
'    MiscTests = testCount
End Function

Sub runRecurIterTestsRefactor(ByVal fluentInput As Variant)
    Dim recurCount1 As Long, iterCount1 As Long
    Dim recurCount2 As Long, iterCount2 As Long
    
    Call validateRecurIterFluentOfsRefactor(fluentInput, tfRecur, tfIter, "main")
    
    Call validateRecurIterFuncCounts2Refactor(tfRecur)
    
    recurCount1 = validateRecurIterFuncCountsRefactor(tfRecur)
    iterCount1 = validateRecurIterFuncCountsRefactor(tfIter)
    Debug.Assert recurCount1 = iterCount1
    
    recurCount2 = validateRecurIterFuncCounts2Refactor(tfRecur)
    iterCount2 = validateRecurIterFuncCounts2Refactor(tfIter)
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

Sub validateRecurIterFluentOfsRefactor(ByVal fluentInput As Variant, ByVal tfRecur As IFluentOf, ByVal tfIter As IFluentOf, ByVal recurIterFuncName As String)
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
    
    For Each test In fluentInput.Meta.tests
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

Sub validateTestsRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant)
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
    
    With fluentInput.Meta
        For Each test In .tests
            If Not VBA.Information.IsNull(test.result) And Not VBA.Information.IsNull(.tests(i).result) Then
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
            End If
        Next test
    End With
End Sub

Private Function getShouldOrShouldNotFromFluentInputAndSetTestingValue(fluentInput As Variant, testingValue As Variant, isShould As Boolean) As IShould
    Dim shouldOrShouldNot  As IShould
    Dim tempFluent As IFluent
    Dim tempFluentOf As IFluentOf
    Dim tempFluentFunction As IFluentFunction
    
    If TypeOf fluentInput Is cFluent Or TypeOf fluentInput Is IFluent Then
        Set tempFluent = fluentInput
        
        If VBA.Information.IsObject(testingValue) Then
            Set tempFluent.testValue = testingValue
        Else
            tempFluent.testValue = testingValue
        End If
        
        If isShould Then
            Set shouldOrShouldNot = tempFluent.Should
        Else
            Set shouldOrShouldNot = tempFluent.ShouldNot
        End If
    ElseIf TypeOf fluentInput Is cFluentOf Or TypeOf fluentInput Is IFluentOf Then
        Set tempFluentOf = fluentInput
        
        If isShould Then
            Set shouldOrShouldNot = tempFluentOf.Of(testingValue).Should
        Else
            Set shouldOrShouldNot = tempFluentOf.Of(testingValue).ShouldNot
        End If
    ElseIf TypeOf fluentInput Is IFluentFunction Or TypeOf fluentInput Is cFluentFunction Then
        Set tempFluentFunction = fluentInput
        
        If Not IsMissing(testingValue) Then
            If isShould Then
                Set shouldOrShouldNot = tempFluentFunction.OfCalledFunction("getVal", testingValue).Should
            Else
                Set shouldOrShouldNot = tempFluentFunction.OfCalledFunction("getVal", testingValue).ShouldNot
            End If
        Else
            If isShould Then
                Set shouldOrShouldNot = tempFluentFunction.OfCalledFunction("getVal", testingValue).Should
            Else
                Set shouldOrShouldNot = tempFluentFunction.OfCalledFunction("getVal", testingValue).ShouldNot
            End If
        End If
    End If
    
    Set getShouldOrShouldNotFromFluentInputAndSetTestingValue = shouldOrShouldNot
End Function

Private Function fluentTester( _
    ByVal fluentInput As Variant, _
    ByVal functionName As String, _
    ByVal isShould As Boolean, _
    Optional ByVal testingValue As Variant, Optional ByVal testingInput As Variant, _
    Optional ByVal testingInput1 As Variant, Optional ByVal testingInput2 As Variant, _
    Optional ByVal testingInput3 As Variant, Optional ByVal testingInput4 As Variant, _
    Optional ByVal testingInput5 As Variant, Optional ByVal testingInput6 As Variant, _
    Optional ByVal testingInput7 As Variant, Optional ByVal testingInput8 As Variant, _
    Optional ByVal testingInput9 As Variant, Optional ByVal testingInput10 As Variant, _
    Optional ByVal testingInput11 As Variant, Optional ByVal testingInput12 As Variant, _
    Optional ByVal testingInput13 As Variant, Optional ByVal testingInput14 As Variant, _
    Optional ByVal testingInput15 As Variant, Optional ByVal testingInput16 As Variant, _
    Optional ByVal testingInput17 As Variant, Optional ByVal testingInput18 As Variant, _
    Optional ByVal testingInput19 As Variant, Optional ByVal testingInput20 As Variant, _
    Optional ByVal testingInput21 As Variant, Optional ByVal testingInput22 As Variant, _
    Optional ByVal testingInput23 As Variant, Optional ByVal testingInput24 As Variant, _
    Optional ByVal testingInput25 As Variant, Optional ByVal testingInput26 As Variant, _
    Optional ByVal testingInput27 As Variant, Optional ByVal testingInput28 As Variant, _
    Optional ByVal testingInput29 As Variant, Optional ByVal testingInput30 As Variant, _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant _
) As Variant
    Dim b As Variant
    Dim ucaseFunctionName As String
    Dim fluentInputShould As IShould
    
    ucaseFunctionName = VBA.Strings.UCase$(functionName)
    b = False
    Set fluentInputShould = getShouldOrShouldNotFromFluentInputAndSetTestingValue(fluentInput, testingValue, isShould)
    
    Select Case ucaseFunctionName
        Case UCase("Alphabetic")
            Call fluentInputShould.Be.Alphabetic
        Case UCase("Alphanumeric")
            Call fluentInputShould.Be.Alphanumeric
        Case UCase("Between")
            Call fluentInputShould.Be.Between(lowerVal, higherVal)
        Case UCase("Contain")
            Call fluentInputShould.Contain(testingInput)
        Case UCase("DepthCountOf")
            Call fluentInputShould.Have.DepthCountOf(testingInput)
        Case UCase("Elements")
            '//passes in all 30 test input parameters and filters out any that have values of missing in the testing function.
            If IsMissing(testingInput) Then
                Call fluentInputShould.Have.Elements(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            Else
                Call fluentInputShould.Have.Elements(testingInput)
            End If
        Case UCase("ElementsInDataStructure")
            Call fluentInputShould.Have.ElementsInDataStructure(testingInput)
        Case UCase("EndWith")
            Call fluentInputShould.EndWith(testingInput)
        Case UCase("Erroneous")
            Call fluentInputShould.Be.Erroneous
        Case UCase("ErrorDescriptionOf")
            Call fluentInputShould.Have.ErrorDescriptionOf(testingInput)
        Case UCase("ErrorNumberOf")
            Call fluentInputShould.Have.ErrorNumberOf(testingInput)
        Case UCase("EqualTo")
            Call fluentInputShould.Be.EqualTo(testingInput)
        Case UCase("EvaluateTo")
            Call fluentInputShould.EvaluateTo(testingInput)
        Case UCase("ExactSameElementsAs")
            Call fluentInputShould.Have.ExactSameElementsAs(testingInput)
        Case UCase("GreaterThan")
            Call fluentInputShould.Be.GreaterThan(testingInput)
        Case UCase("GreaterThanOrEqualTo")
            Call fluentInputShould.Be.GreaterThanOrEqualTo(testingInput)
        Case UCase("IdenticalTo")
            Call fluentInputShould.Be.IdenticalTo(testingInput)
        Case UCase("InDataStructure")
            Call fluentInputShould.Be.InDataStructure(testingInput)
            '//passes in all 30 test input parameters and filters out any that have values of missing in the testing function.
        Case UCase("InDataStructures")
            If IsMissing(testingInput) Then
                Call fluentInputShould.Be.InDataStructures(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            Else
                Call fluentInputShould.Be.InDataStructures(testingInput)
            End If
        Case UCase("LengthBetween")
            Call fluentInputShould.Have.LengthBetween(lowerVal, higherVal)
        Case UCase("LengthOf")
            Call fluentInputShould.Have.LengthOf(testingInput)
        Case UCase("LessThan")
            Call fluentInputShould.Be.LessThan(testingInput)
        Case UCase("LessThanOrEqualTo")
            Call fluentInputShould.Be.LessThanOrEqualTo(testingInput)
        Case UCase("MaxLengthOf")
            Call fluentInputShould.Have.MaxLengthOf(testingInput)
        Case UCase("MinLengthOf")
            Call fluentInputShould.Have.MinLengthOf(testingInput)
        Case UCase("NestedCountOf")
            Call fluentInputShould.Have.NestedCountOf(testingInput)
        Case UCase("Numeric")
            Call fluentInputShould.Be.Numeric
            '//passes in all 30 test input parameters and filters out any that have values of missing in the testing function.
        Case UCase("OneOf")
            If IsMissing(testingInput) Then
                Call fluentInputShould.Be.OneOf(testingInput1, testingInput2, testingInput3, testingInput4, testingInput5, testingInput6, testingInput7, testingInput8, testingInput9, testingInput10, testingInput11, testingInput12, testingInput13, testingInput14, testingInput15, testingInput16, testingInput17, testingInput18, testingInput19, testingInput20, testingInput21, testingInput22, testingInput23, testingInput24, testingInput25, testingInput26, testingInput27, testingInput28, testingInput29, testingInput30)
            Else
                Call fluentInputShould.Be.OneOf(testingInput)
            End If
        Case UCase("Procedure")
            Call fluentInputShould.Have.Procedure(testingInput1, testingInput2)
        Case UCase("SameElementsAs")
            Call fluentInputShould.Have.SameElementsAs(testingInput)
        Case UCase("SameTypeAs")
            Call fluentInputShould.Have.SameTypeAs(testingInput)
        Case UCase("SameUniqueElementsAs")
            Call fluentInputShould.Have.SameUniqueElementsAs(testingInput)
        Case UCase("Something")
            Call fluentInputShould.Be.Something
        Case UCase("StartWith")
            Call fluentInputShould.StartWith(testingInput)
        Case Else
            MsgBox "method " & ucaseFunctionName & " not supported", vbCritical
    End Select
    
    Set fluentTester = fluentInput
End Function

Private Sub printTestCountRefactor(ByVal testCount As Long)
    If testCount > 1 Then
        Debug.Print testCount & " tests finished!" & vbNewLine
    ElseIf testCount = 1 Then
        Debug.Print "1 Test finished!"
    End If
End Sub

Private Function InitFluentInput(fluentInput As Variant) As Variant
    With fluentInput.Meta
        .Printing.PassedMessage = "Success"
        .Printing.FailedMessage = "Failure"
        .Printing.UnexpectedMessage = "What?"
        
        .tests.ToStrDev = True
    End With
    
    
    Set InitFluentInput = fluentInput
End Function

Function validateRecurIterFuncCountsRefactor(ByVal recurIterFluentOf As cFluentOf) As Long
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

    validateRecurIterFuncCountsRefactor = counter
End Function

Function validateRecurIterFuncCounts2Refactor(ByVal recurIterFluentOf As cFluentOf) As Long
    Dim TestingInfoDev As ITestingFunctionsInfoDev
    Dim counter As Long
    Dim recurIterFuncNameCol As VBA.Collection
    Dim elem As Variant
    Dim testSubInfoRecur As ITestingFunctionsInfo
    Dim testSubInfoIter As ITestingFunctionsInfo

    'In order for this function to work correctly, the individual testing functions
    'for the recursive and iterative methods must be called.

    Set TestingInfoDev = recurIterFluentOf.Meta.tests.TestingFunctionsInfos
    Set recurIterFuncNameCol = TestingInfoDev.getRecurIterFuncNameCol
    counter = 0

    For Each elem In recurIterFuncNameCol
        Set testSubInfoRecur = VBA.Interaction.CallByName(TestingInfoDev, elem & "Recur", VbGet)
        Set testSubInfoIter = VBA.Interaction.CallByName(TestingInfoDev, elem & "Iter", VbGet)

        Debug.Assert testSubInfoIter.Count = testSubInfoRecur.Count

        If testSubInfoIter.Count = testSubInfoRecur.Count Then counter = counter + 1
    Next elem

    validateRecurIterFuncCounts2Refactor = counter
End Function

Private Function getAndInitEventRefactor(ByVal fluent As IFluent, ByVal fluentInput As Variant, ByVal testFluentResult As IFluentOf) As zEvents
    Set mEvents = New zEvents
    
    Set mEvents.setFluent = fluent
    
    If TypeOf fluentInput Is cFluentOf Or TypeOf fluentInput Is IFluentOf Then
        Set mEvents.setFluentOfTest = fluentInput
    ElseIf TypeOf fluentInput Is cFluent Or TypeOf fluentInput Is IFluent Then
        Set mEvents.setFluentTest = fluentInput
    ElseIf TypeOf fluentInput Is cFluentFunction Or TypeOf fluentInput Is IFluentFunction Then
        Set mEvents.setFluentFunctionTest = fluentInput
    End If
    
    Set mEvents.setFluentEventOfResult = testFluentResult
    
    Set getAndInitEventRefactor = mEvents
End Function

Public Function getVal(Optional val As Variant) As Variant
    If VBA.Information.IsObject(val) Then
        Set getVal = val
    Else
        getVal = val
    End If
End Function
