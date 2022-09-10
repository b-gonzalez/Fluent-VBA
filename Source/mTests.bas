Attribute VB_Name = "mTests"
Option Explicit

Private Sub FluentAAAExamples()
    Dim Result1 As cFluent
    Dim Result2 As cFluentOf
    Dim returnedResult As Variant
    
    '//arrange
    Set Result1 = New cFluent
    Set Result2 = New cFluentOf
    returnedResult = returnVal(5)
    
    '//Act
    Result1.TestValue = 5 + 0
    
    '//Assert
    With Result1.Should.Be
        Debug.Assert .EqualTo(5)
    End With
    

    
    '//Act
    With Result2.Of(returnedResult).Should
        '//Assert
         .Be.EqualTo (6)
    End With
    
    '//Act & Assert
    Debug.Assert Result2.Of(returnedResult).Should.Be.EqualTo(5)
    
    Result1.Meta.Printing.PrintToImmediate
    Result2.Meta.Printing.PrintToImmediate
End Sub

Private Function returnVal(value As Variant)
    returnVal = value
End Function

Public Sub runMainTests()
    Dim fluent As cFluent
    Dim testFluent As cFluentOf
    
    Set fluent = New cFluent
    Set testFluent = New cFluentOf
    
    With fluent.Meta.Printing
        .TestName = "Result"
    End With
    
    fluent.Meta.Printing.Category = "metaTests"
    testFluent.Meta.Printing.Category = "metaTests"
    Call MetaTests(fluent, testFluent)
    
    fluent.Meta.Printing.Category = "EqualityTests"
    testFluent.Meta.Printing.Category = "EqualityTests"
    Call EqualityTests(fluent, testFluent)

    fluent.Meta.Printing.Category = "positiveDocumentationTests"
    testFluent.Meta.Printing.Category = "positiveDocumentationTests"
    Call positiveDocumentationTests(fluent, testFluent)

    fluent.Meta.Printing.Category = "negativeDocumentationTests"
    testFluent.Meta.Printing.Category = "negativeDocumentationTests"
    Call negativeDocumentationTests(fluent, testFluent)

    Debug.Print "All tests Finished!"
    Call printTestCount(testFluent.Meta.TestCount)
    
    fluent.Meta.Printing.PrintToSheet
    testFluent.Meta.Printing.PrintToSheet
    'Fluent.Meta.Printing.PrintToImmediate
End Sub

Public Sub runExamples()
    Dim fluent As cFluent
    
    Set fluent = New cFluent
    
    Call Example1(fluent)
    
    'Call Fluent.Meta.Printing.PrintToSheet
    'Call fluent.Meta.Printing.PrintToImmediate
End Sub

Private Sub printTestCount(TestCount As Long)
    If TestCount > 1 Then
        Debug.Print TestCount & " tests finished!"
    ElseIf TestCount = 1 Then
        Debug.Print "Test finished!"
    End If
End Sub

Private Sub Example1(result As cFluent)
    result.Meta.Printing.Category = "Example 1"
    result.TestValue = 5 + 5

    result.Should.Be.EqualTo (10) 'true
    result.Should.Be.GreaterThan (9) 'true
    result.Should.Be.LessThan (11) 'true
    result.ShouldNot.Be.EqualTo (9) 'true
    result.ShouldNot.Contain (4) 'true
    result.Should.StartWith (1) 'true
    result.Should.EndWith (0) 'true
    result.Should.Contain (10) 'true
    
    result.Should.EndWith (9) 'false
    result.ShouldNot.StartWith (1) 'false
    result.ShouldNot.EndWith (0) 'false
    
    result.ShouldNot.Have.LengthOf (0) 'true
    result.ShouldNot.Have.MaxLengthOf (0) 'true
    result.ShouldNot.Have.MinLengthOf (3) 'true

    result.Should.Have.LengthOf (0) 'false
    result.Should.Have.MaxLengthOf (1) 'false
    result.Should.Have.MinLengthOf (3) 'false
    
    Debug.Print result.Meta.TestCount & " tests finished"
    
End Sub

Private Sub Example2()
    Dim testNums As Long
    Dim result As cFluent
    Dim TestNames() As String
    Dim i As Long
    Dim temp As Boolean
    
    Set result = New cFluent
    result.TestValue = 10
    
    With result
        Debug.Assert .Should.Be.EqualTo(10) And .Should.Be.GreaterThan(0) 'true
        Debug.Assert .Should.Be.EqualTo(10) And .Should.Be.GreaterThan(0) And .Should.Have.LengthOf(2) 'true
        
        Debug.Assert .Should.Be.EqualTo(10) Or .Should.Be.GreaterThan(0) 'true
        Debug.Assert .Should.Be.EqualTo(10) Or .Should.Be.GreaterThan(0) Or .Should.Have.LengthOf(2) 'true
        
        Debug.Assert .Should.Be.EqualTo(10) And .Should.Be.GreaterThan(0) Or .Should.Have.LengthOf(2) 'true
        Debug.Assert .Should.Be.EqualTo(10) Or .Should.Be.GreaterThan(0) And .Should.Have.LengthOf(2) 'true
        
        Debug.Assert .Should.Be.EqualTo(10) And .Should.Be.GreaterThan(11) 'false
        Debug.Assert .Should.Be.EqualTo(9) Or .Should.Be.GreaterThan(11) 'false
    End With
End Sub

Private Sub Example3()
    Dim testNums As Long
    Dim result As cFluent
    Dim TestNames() As String
    Dim i As Long
    'Dim testResults(4) As Boolean
    Dim temp As Boolean
    
    Set result = New cFluent
    result.TestValue = 10
    
    With result
        .Meta.Printing.TestName = "Test - Result should be equal to 10 - "
        Debug.Assert .Should.Be.EqualTo(10)  ' true
        
        .Meta.Printing.TestName = "Test - Result should greater than 9 - "
        Debug.Assert .Should.Be.GreaterThan(9)  'true
        
        .Meta.Printing.TestName = "Test - Result should be less than 11 - "
        Debug.Assert .Should.Be.LessThan(11)  ' true
        
        .Meta.Printing.TestName = "Test - Result should not be equal to 9 - "
        Debug.Assert .ShouldNot.Be.EqualTo(9)  'true
        
        .Meta.Printing.TestName = "Test - Result should not contain 4 - "
        Debug.Assert .ShouldNot.Contain(4)  'true
        
        .Meta.Printing.TestName = "Test - Result should start with 1 - "
        Debug.Assert .Should.StartWith(1)  'true
        
        .Meta.Printing.TestName = "Test - Result should end with 0 - "
        Debug.Assert .Should.EndWith(0)  'true
    
        .Meta.Printing.TestName = "Test - Result should contain 10 - "
        Debug.Assert .Should.Contain(10)  'true
    
        .Meta.Printing.TestName = "Test - Result should end with 9 - "
        Debug.Assert .Should.EndWith(9)  'false
        
        .Meta.Printing.TestName = "Test -  - "
        Debug.Assert .ShouldNot.StartWith(1)  'false
        
        .Meta.Printing.TestName = "Test - Result shoudl not end with 0  - "
        Debug.Assert .ShouldNot.EndWith(0)  'false
        
        .Meta.Printing.TestName = "Test - result should not have length of 0 - "
        Debug.Assert .ShouldNot.Have.LengthOf(0)  'true
        
        .Meta.Printing.TestName = "Test - result should not have max length of 0 - "
        Debug.Assert .ShouldNot.Have.MaxLengthOf(0)  'true
        
        .Meta.Printing.TestName = "Test - result should not have min length of 3 - "
        Debug.Assert .ShouldNot.Have.MinLengthOf(3)  'true
        
        .Meta.Printing.TestName = "Test - result should have length of 0 - "
        Debug.Assert .Should.Have.LengthOf(0)  'false
        
        .Meta.Printing.TestName = "Test - result should have max length of 1 - "
        Debug.Assert .Should.Have.MaxLengthOf(1)  'false
        
        .Meta.Printing.TestName = "Test - result should have min length of 3 - "
        Debug.Assert .Should.Have.MinLengthOf(3)  'false
    End With
End Sub

Private Sub Example4()
    Dim testNums As Long
    Dim result() As cFluent
    Dim TestNames() As String
    Dim i As Long
    Dim testResults() As Boolean
    Dim temp As Boolean
    
    testNums = 16
    
    ReDim result(testNums)
    ReDim TestNames(testNums)
    ReDim testResults(testNums)
    
    TestNames(0) = "Test - Result should be equal to 10 - "
    TestNames(1) = "Test - Result should greater than 9 - "
    TestNames(2) = "Test - Result should be less than 11 - "
    TestNames(3) = "Test - Result should not be equal to 9 - "
    TestNames(4) = "Test - Result should not contain 4 - "
    TestNames(5) = "Test - Result should start with 1 - "
    TestNames(6) = "Test - Result should end with 0 - "
    TestNames(7) = "Test - Result should contain 10 - "
    TestNames(8) = "Test - Result should end with 9 - "
    TestNames(9) = "Test - Result should not start with 1 - "
    TestNames(10) = "Test - Result should not end with 0 - "
    TestNames(11) = "Test - Result should not have length of 0 - "
    TestNames(12) = "Test - Result should not have max length of 0 - "
    TestNames(13) = "Test - Result should not have min length of 3 - "
    TestNames(14) = "Test - Result should have length of 0 - "
    TestNames(15) = "Test - Result should have max length of 1 - "
    TestNames(16) = "Test - Result should have have min length of 3 - "
    
    For i = LBound(result) To UBound(result)
        Set result(i) = New cFluent
        result(i).Meta.Printing.TestName = TestNames(i)
        'Result(i).Meta.PrintSettings.PrintTestsToImmediate = True
        result(i).TestValue = 10
    Next i
    
    Debug.Assert result(0).Should.Be.EqualTo(10) 'true
    Debug.Assert result(1).Should.Be.GreaterThan(9) 'true
    Debug.Assert result(2).Should.Be.LessThan(11) 'true
    Debug.Assert result(3).ShouldNot.Be.EqualTo(9) 'true
    Debug.Assert result(4).ShouldNot.Contain(4) 'true
    Debug.Assert result(5).Should.StartWith(1) 'true
    Debug.Assert result(6).Should.EndWith(0) 'true
    Debug.Assert result(7).Should.Contain(10) 'trues
    Debug.Assert result(8).Should.EndWith(9) 'false
    Debug.Assert result(9).ShouldNot.StartWith(1) 'false
    Debug.Assert result(10).ShouldNot.EndWith(0) 'false
    Debug.Assert result(11).ShouldNot.Have.LengthOf(0) 'true
    Debug.Assert result(12).ShouldNot.Have.MaxLengthOf(0) 'true
    Debug.Assert result(13).ShouldNot.Have.MinLengthOf(3) 'true
    Debug.Assert result(14).Should.Have.LengthOf(0) 'false
    Debug.Assert result(15).Should.Have.MaxLengthOf(1) 'false
    Debug.Assert result(16).Should.Have.MinLengthOf(3) 'false
End Sub

Private Sub Example5()
    Dim testNums As Long
    Dim result() As cFluent
    Dim TestNames() As String
    Dim i As Long
    Dim testResults() As Boolean
    Dim temp As Boolean
    
    testNums = 16
    
    ReDim result(testNums)
    ReDim TestNames(testNums)
    ReDim testResults(testNums)
    
    TestNames(0) = "Test - Result should be equal to 10 - "
    TestNames(1) = "Test - Result should greater than 9 - "
    TestNames(2) = "Test - Result should be less than 11 - "
    TestNames(3) = "Test - Result should not be equal to 9 - "
    TestNames(4) = "Test - Result should not contain 4 - "
    TestNames(5) = "Test - Result should start with 1 - "
    TestNames(6) = "Test - Result should end with 0 - "
    TestNames(7) = "Test - Result should contain 10 - "
    TestNames(8) = "Test - Result should end with 9 - "
    TestNames(9) = "Test - Result should not start with 1 - "
    TestNames(10) = "Test - Result should not end with 0 - "
    TestNames(11) = "Test - Result should not have length of 0 - "
    TestNames(12) = "Test - Result should not have max length of 0 - "
    TestNames(13) = "Test - Result should not have min length of 3 - "
    TestNames(14) = "Test - Result should have length of 0 - "
    TestNames(15) = "Test - Result should have max length of 1 - "
    TestNames(16) = "Test - Result should have have min length of 3 - "
    
    For i = LBound(result) To UBound(result)
        Set result(i) = New cFluent
        result(i).Meta.Printing.TestName = TestNames(i)
        result(i).TestValue = 10
    Next i
    
    testResults(0) = result(0).Should.Be.EqualTo(10) 'true
    testResults(1) = result(1).Should.Be.GreaterThan(9) 'true
    testResults(2) = result(2).Should.Be.LessThan(11) 'true
    testResults(3) = result(3).ShouldNot.Be.EqualTo(9) 'true
    testResults(4) = result(4).ShouldNot.Contain(4) 'true
    testResults(5) = result(5).Should.StartWith(1) 'true
    testResults(6) = result(6).Should.EndWith(0) 'true
    testResults(7) = result(7).Should.Contain(10) 'true
    testResults(8) = result(8).Should.EndWith(9) 'false
    testResults(9) = result(9).ShouldNot.StartWith(1) 'false
    testResults(10) = result(10).ShouldNot.EndWith(0) 'false
    testResults(11) = result(11).ShouldNot.Have.LengthOf(0) 'true
    testResults(12) = result(12).ShouldNot.Have.MaxLengthOf(0) 'true
    testResults(13) = result(13).ShouldNot.Have.MinLengthOf(3) 'true
    testResults(14) = result(14).Should.Have.LengthOf(0) 'false
    testResults(15) = result(15).Should.Have.MaxLengthOf(1) 'false
    testResults(16) = result(16).Should.Have.MinLengthOf(3) 'false
    
    For i = LBound(testResults) To UBound(testResults)
        temp = testResults(i)
        Debug.Print temp
        Debug.Assert temp
    Next i
End Sub

Private Sub Example6()
    Dim testNums As Long
    Dim result As cFluent
    Dim TestNames() As String
    Dim i As Long
    'Dim testResults(4) As Boolean
    Dim temp As Boolean
    
    Set result = New cFluent
    result.TestValue = 10
    
    result.Meta.Printing.TestName = "Test - Result should be equal to 10 - "
    Debug.Assert result.Should.Be.EqualTo(10)  ' true
    
    result.Meta.Printing.TestName = "Test - Result should greater than 9 - "
    Debug.Assert result.Should.Be.GreaterThan(9)  'true
    
    result.Meta.Printing.TestName = "Test - Result should be less than 11 - "
    Debug.Assert result.Should.Be.LessThan(11)  ' true
    
    result.Meta.Printing.TestName = "Test - Result should not be equal to 9 - "
    Debug.Assert result.ShouldNot.Be.EqualTo(9)   'true
    
    result.Meta.Printing.TestName = "Test - Result should not contain 4 - "
    Debug.Assert result.ShouldNot.Contain(4)  'true
    
    result.Meta.Printing.TestName = "Test - Result should start with 1 - "
    Debug.Assert result.Should.StartWith(1)  'true
    
    result.Meta.Printing.TestName = "Test - Result should end with 0 - "
    Debug.Assert result.Should.EndWith(0)  'true

    result.Meta.Printing.TestName = "Test - Result should contain 10 - "
    Debug.Assert result.Should.Contain(10)  'true

    result.Meta.Printing.TestName = "Test - Result should end with 9 - "
    Debug.Assert result.Should.EndWith(9)  'false
    
    result.Meta.Printing.TestName = "Test -  - "
    Debug.Assert result.ShouldNot.StartWith(1)  'false
    
    result.Meta.Printing.TestName = "Test - Result shoudl not end with 0  - "
    Debug.Assert result.ShouldNot.EndWith(0)  'false
    
    result.Meta.Printing.TestName = "Test - result should not have length of 0 - "
    Debug.Assert result.ShouldNot.Have.LengthOf(0)  'true
    
    result.Meta.Printing.TestName = "Test - result should not have max length of 0 - "
    Debug.Assert result.ShouldNot.Have.MaxLengthOf(0)  'true
    
    result.Meta.Printing.TestName = "Test - result should not have min length of 3 - "
    Debug.Assert result.ShouldNot.Have.MinLengthOf(3)  'true
    
    result.Meta.Printing.TestName = "Test - result should have length of 0 - "
    Debug.Assert result.Should.Have.LengthOf(0)  'false
    
    result.Meta.Printing.TestName = "Test - result should have max length of 1 - "
    Debug.Assert result.Should.Have.MaxLengthOf(1)  'false
    
    result.Meta.Printing.TestName = "Test - result should have min length of 3 - "
    Debug.Assert result.Should.Have.MinLengthOf(3)  'false
    
End Sub

Private Sub MetaTests(fluent As cFluent, testFluent As cFluentOf)
    Call checkTestCount(fluent, testFluent)
End Sub

Private Sub checkTestCount(testFluent As cFluent, fluent As cFluentOf)
    Dim result As Variant
    
    result = testFluent.Meta.TestCount
    Debug.Assert fluent.Of(result).Should.Be.EqualTo(0)
    
    testFluent.TestValue = "Test"
    Debug.Assert testFluent.Should.Be.EqualTo("Test")
    result = testFluent.Meta.TestCount
    Debug.Assert fluent.Of(result).Should.Be.EqualTo(1)
    
    testFluent.TestValue = "Test"
    Debug.Assert testFluent.Should.Be.EqualTo("Test")
    result = testFluent.Meta.TestCount
    Debug.Assert fluent.Of(result).Should.Be.EqualTo(2)
End Sub

Private Sub EqualityTests(fluent As cFluent, testFluent As cFluentOf)
    Dim TestResult As Boolean
    
    fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(False).Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)

    fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)

    fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    '//Approximate equality tests

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
End Sub

Private Sub positiveDocumentationTests(fluent As cFluent, testFluent As cFluentOf)
    'Dim testFluent As cFluentOf
    Dim TestResult As Boolean
    Dim col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    
    'Set testFluent = New cFluentOf
    
    fluent.TestValue = testFluent.Of(10).Should.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Contain(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Contain(0)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Contain(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Contain(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.StartWith(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.StartWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.EndWith(0)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.EndWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Have.MinLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.Between(9, 11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(1, 3)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(9, 10, 11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ' //Object and data structure tests
    
    Set col = New Collection
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(col).Should.Be.OneOf(col, d)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(col, d, 10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(col).Should.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = Nothing
    fluent.TestValue = testFluent.Of(col).Should.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    arr = Array(9, 10, 11)
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ReDim arr(1, 1)
    arr(0, 0) = 9
    arr(0, 1) = 10
    arr(1, 0) = 11
    arr(1, 1) = 12
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ReDim arr(1, 1, 1)
    arr(0, 0, 0) = 6
    arr(0, 0, 1) = 7
    arr(0, 1, 0) = 8
    arr(0, 1, 1) = 9
    arr(1, 0, 0) = 10
    arr(1, 0, 1) = 11
    arr(1, 1, 0) = 12
    arr(1, 1, 1) = 13
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    arr = Array(9, Array(10, Array(11)))
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = New Collection
    col.Add 9
    col.Add 10
    col.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = New Collection
    col.Add 9
    col.Add Array(10, Array(11))
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    Set col = Nothing
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, Array(10, Array(11))
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 9, 1
    d.Add 10, 2
    d.Add 11, 3
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.keys)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set al = CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ' //Approximate equality tests
    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("10").Should.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("True").Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of(5.0000001).Should.Be.EqualTo(5)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    testFluent.Meta.ApproximateEqual = False
    
    '//Evaluation tests
    
    fluent.TestValue = testFluent.Of(True).Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(True).Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(False).Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(False).Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("true").Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("false").Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("TRUE").Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("FALSE").Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(-1).Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(-1).Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(0).Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(0).Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("-1").Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("-1").Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("0").Should.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("0").Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(5 + 5).Should.EvaluateTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("5 + 5").Should.EvaluateTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("5 + 5 = 10").Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("5 + 5 > 9").Should.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    '//Testing errors is possible if they're put in strings
    fluent.TestValue = testFluent.Of("1 / 0").Should.EvaluateTo(CVErr(xlErrDiv0))
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
End Sub

Private Sub negativeDocumentationTests(fluent As cFluent, testFluent As cFluentOf)
    'Dim testFluent As cFluentOf
    Dim TestResult As Boolean
    Dim col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    
    'Set testFluent = New cFluentOf
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(0)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(0)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MinLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(9, 11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(1, 3)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(9, 10, 11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ' //Object and data structure tests
    
    Set col = New Collection
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.OneOf(col, d)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(col, d, 10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set col = Nothing
    fluent.TestValue = testFluent.Of(col).ShouldNot.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    arr = Array(9, 10, 11)
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ReDim arr(1, 1)
    arr(0, 0) = 9
    arr(0, 1) = 10
    arr(1, 0) = 11
    arr(1, 1) = 12
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ReDim arr(1, 1, 1)
    arr(0, 0, 0) = 6
    arr(0, 0, 1) = 7
    arr(0, 1, 0) = 8
    arr(0, 1, 1) = 9
    arr(1, 0, 0) = 10
    arr(1, 0, 1) = 11
    arr(1, 1, 0) = 12
    arr(1, 1, 1) = 13
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    arr = Array(9, Array(10, Array(11)))
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set col = New Collection
    col.Add 9
    col.Add 10
    col.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set col = New Collection
    col.Add 9
    col.Add Array(10, Array(11))
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    Set col = Nothing
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, Array(10, Array(11))
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 9, 1
    d.Add 10, 2
    d.Add 11, 3
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d.keys)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set al = CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(al)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ' //Approximate equality tests
    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("10").ShouldNot.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of("True").ShouldNot.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    testFluent.Meta.ApproximateEqual = True
    fluent.TestValue = testFluent.Of(5.0000001).ShouldNot.Be.EqualTo(5)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    testFluent.Meta.ApproximateEqual = False
    
    '//Evaluation tests
    
    fluent.TestValue = testFluent.Of(True).ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(True).ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(False).ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(False).ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("true").ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("false").ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("TRUE").ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("FALSE").ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(-1).ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(-1).ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(0).ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of(0).ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("-1").ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("-1").ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of("0").ShouldNot.EvaluateTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("0").ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    fluent.TestValue = testFluent.Of(5 + 5).ShouldNot.EvaluateTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("5 + 5").ShouldNot.EvaluateTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("5 + 5 = 10").ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    fluent.TestValue = testFluent.Of("5 + 5 > 9").ShouldNot.EvaluateTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    '//Testing errors is possible if they're put in strings
    fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.EvaluateTo(CVErr(xlErrDiv0))
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
End Sub
