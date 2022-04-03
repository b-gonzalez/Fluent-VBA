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
    Result1.TestValue = returnedResult
    
    '//Assert
    Debug.Assert Result1.Should.Be.EqualTo(5)
    
    '//Act
    With Result2.Of(returnedResult).Should
        '//Assert
        Debug.Assert .Be.EqualTo(5) And .Be.LessThanOrEqualTo(5) And .Be.GreaterThanOrEqualTo(5)
    End With
    
    '//Act & Assert
    Debug.Assert Result2.Of(returnedResult).Should.Be.EqualTo(5)
End Sub

Private Function returnVal(value As Variant)
    returnVal = value
End Function

Public Sub runMainTests()
    Dim fluent As cFluentOf
    Dim testFluent As cFluent
    
    Set fluent = New cFluentOf
    Set testFluent = New cFluent
    
    testFluent.Meta.PrintTotalTests = True
    
    Call MetaTests(fluent, testFluent)
    Call positiveDocumentationTests(testFluent)
    Call negativeDocumentationTests(testFluent)
    Debug.Print "All tests Finished!"
End Sub

Private Sub Example1()
    Dim Result As cFluent
    Set Result = New cFluent
    Result.TestValue = 10
       
    Result.Meta.PrintResults = True
    
    Result.Should.Be.EqualTo (10) 'true
    Result.Should.Be.GreaterThan (9) 'true
    Result.Should.Be.LessThan (11) 'true
    Result.ShouldNot.Be.EqualTo (9) 'true
    Result.ShouldNot.Contain (4) 'true
    Result.Should.StartWith (1) 'true
    Result.Should.EndWith (0) 'true
    Result.Should.Contain (10) 'true
    Result.Should.EndWith (9) 'false

    Result.ShouldNot.StartWith (1) 'false
    Result.ShouldNot.EndWith (0) 'false
    Result.ShouldNot.Have.LengthOf (0) 'true
    Result.ShouldNot.Have.MaxLengthOf (0) 'true
    Result.ShouldNot.Have.MinLengthOf (3) 'true

    Result.Should.Have.LengthOf (0) 'false
    Result.Should.Have.MaxLengthOf (1) 'false
    Result.Should.Have.MinLengthOf (3) 'false
    
End Sub

Private Sub Example2()
    Dim testNums As Long
    Dim Result As cFluent
    Dim TestNames() As String
    Dim i As Long
    Dim temp As Boolean
    
    Set Result = New cFluent
    Result.TestValue = 10
    
    With Result
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
    Dim Result As cFluent
    Dim TestNames() As String
    Dim i As Long
    'Dim testResults(4) As Boolean
    Dim temp As Boolean
    
    Set Result = New cFluent
    Result.TestValue = 10
    
    With Result
        .Meta.TestName = "Test - Result should be equal to 10 - "
        Debug.Assert .Should.Be.EqualTo(10)  ' true
        
        .Meta.TestName = "Test - Result should greater than 9 - "
        Debug.Assert .Should.Be.GreaterThan(9)  'true
        
        .Meta.TestName = "Test - Result should be less than 11 - "
        Debug.Assert .Should.Be.LessThan(11)  ' true
        
        .Meta.TestName = "Test - Result should not be equal to 9 - "
        Debug.Assert .ShouldNot.Be.EqualTo(9)  'true
        
        .Meta.TestName = "Test - Result should not contain 4 - "
        Debug.Assert .ShouldNot.Contain(4)  'true
        
        .Meta.TestName = "Test - Result should start with 1 - "
        Debug.Assert .Should.StartWith(1)  'true
        
        .Meta.TestName = "Test - Result should end with 0 - "
        Debug.Assert .Should.EndWith(0)  'true
    
        .Meta.TestName = "Test - Result should contain 10 - "
        Debug.Assert .Should.Contain(10)  'true
    
        .Meta.TestName = "Test - Result should end with 9 - "
        Debug.Assert .Should.EndWith(9)  'false
        
        .Meta.TestName = "Test -  - "
        Debug.Assert .ShouldNot.StartWith(1)  'false
        
        .Meta.TestName = "Test - Result shoudl not end with 0  - "
        Debug.Assert .ShouldNot.EndWith(0)  'false
        
        .Meta.TestName = "Test - result should not have length of 0 - "
        Debug.Assert .ShouldNot.Have.LengthOf(0)  'true
        
        .Meta.TestName = "Test - result should not have max length of 0 - "
        Debug.Assert .ShouldNot.Have.MaxLengthOf(0)  'true
        
        .Meta.TestName = "Test - result should not have min length of 3 - "
        Debug.Assert .ShouldNot.Have.MinLengthOf(3)  'true
        
        .Meta.TestName = "Test - result should have length of 0 - "
        Debug.Assert .Should.Have.LengthOf(0)  'false
        
        .Meta.TestName = "Test - result should have max length of 1 - "
        Debug.Assert .Should.Have.MaxLengthOf(1)  'false
        
        .Meta.TestName = "Test - result should have min length of 3 - "
        Debug.Assert .Should.Have.MinLengthOf(3)  'false
    End With
End Sub

Private Sub Example4()
    Dim testNums As Long
    Dim Result() As cFluent
    Dim TestNames() As String
    Dim i As Long
    Dim testResults() As Boolean
    Dim temp As Boolean
    
    testNums = 16
    
    ReDim Result(testNums)
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
    
    For i = LBound(Result) To UBound(Result)
        Set Result(i) = New cFluent
        Result(i).Meta.TestName = TestNames(i)
        Result(i).Meta.PrintResults = True
        Result(i).TestValue = 10
    Next i
    
    Debug.Assert Result(0).Should.Be.EqualTo(10) 'true
    Debug.Assert Result(1).Should.Be.GreaterThan(9) 'true
    Debug.Assert Result(2).Should.Be.LessThan(11) 'true
    Debug.Assert Result(3).ShouldNot.Be.EqualTo(9) 'true
    Debug.Assert Result(4).ShouldNot.Contain(4) 'true
    Debug.Assert Result(5).Should.StartWith(1) 'true
    Debug.Assert Result(6).Should.EndWith(0) 'true
    Debug.Assert Result(7).Should.Contain(10) 'trues
    Debug.Assert Result(8).Should.EndWith(9) 'false
    Debug.Assert Result(9).ShouldNot.StartWith(1) 'false
    Debug.Assert Result(10).ShouldNot.EndWith(0) 'false
    Debug.Assert Result(11).ShouldNot.Have.LengthOf(0) 'true
    Debug.Assert Result(12).ShouldNot.Have.MaxLengthOf(0) 'true
    Debug.Assert Result(13).ShouldNot.Have.MinLengthOf(3) 'true
    Debug.Assert Result(14).Should.Have.LengthOf(0) 'false
    Debug.Assert Result(15).Should.Have.MaxLengthOf(1) 'false
    Debug.Assert Result(16).Should.Have.MinLengthOf(3) 'false
End Sub

Private Sub Example5()
    Dim testNums As Long
    Dim Result() As cFluent
    Dim TestNames() As String
    Dim i As Long
    Dim testResults() As Boolean
    Dim temp As Boolean
    
    testNums = 16
    
    ReDim Result(testNums)
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
    
    For i = LBound(Result) To UBound(Result)
        Set Result(i) = New cFluent
        Result(i).Meta.TestName = TestNames(i)
        Result(i).TestValue = 10
    Next i
    
    testResults(0) = Result(0).Should.Be.EqualTo(10) 'true
    testResults(1) = Result(1).Should.Be.GreaterThan(9) 'true
    testResults(2) = Result(2).Should.Be.LessThan(11) 'true
    testResults(3) = Result(3).ShouldNot.Be.EqualTo(9) 'true
    testResults(4) = Result(4).ShouldNot.Contain(4) 'true
    testResults(5) = Result(5).Should.StartWith(1) 'true
    testResults(6) = Result(6).Should.EndWith(0) 'true
    testResults(7) = Result(7).Should.Contain(10) 'true
    testResults(8) = Result(8).Should.EndWith(9) 'false
    testResults(9) = Result(9).ShouldNot.StartWith(1) 'false
    testResults(10) = Result(10).ShouldNot.EndWith(0) 'false
    testResults(11) = Result(11).ShouldNot.Have.LengthOf(0) 'true
    testResults(12) = Result(12).ShouldNot.Have.MaxLengthOf(0) 'true
    testResults(13) = Result(13).ShouldNot.Have.MinLengthOf(3) 'true
    testResults(14) = Result(14).Should.Have.LengthOf(0) 'false
    testResults(15) = Result(15).Should.Have.MaxLengthOf(1) 'false
    testResults(16) = Result(16).Should.Have.MinLengthOf(3) 'false

    
    For i = LBound(testResults) To UBound(testResults)
        temp = testResults(i)
        Debug.Assert temp
        Debug.Print temp
    Next i
End Sub

Private Sub Example6()
    Dim testNums As Long
    Dim Result As cFluent
    Dim TestNames() As String
    Dim i As Long
    'Dim testResults(4) As Boolean
    Dim temp As Boolean
    
    Set Result = New cFluent
    Result.TestValue = 10
    
    Result.Meta.TestName = "Test - Result should be equal to 10 - "
    Debug.Assert Result.Should.Be.EqualTo(10)  ' true
    
    Result.Meta.TestName = "Test - Result should greater than 9 - "
    Debug.Assert Result.Should.Be.GreaterThan(9)  'true
    
    Result.Meta.TestName = "Test - Result should be less than 11 - "
    Debug.Assert Result.Should.Be.LessThan(11)  ' true
    
    Result.Meta.TestName = "Test - Result should not be equal to 9 - "
    Debug.Assert Result.ShouldNot.Be.EqualTo(9)   'true
    
    Result.Meta.TestName = "Test - Result should not contain 4 - "
    Debug.Assert Result.ShouldNot.Contain(4)  'true
    
    Result.Meta.TestName = "Test - Result should start with 1 - "
    Debug.Assert Result.Should.StartWith(1)  'true
    
    Result.Meta.TestName = "Test - Result should end with 0 - "
    Debug.Assert Result.Should.EndWith(0)  'true

    Result.Meta.TestName = "Test - Result should contain 10 - "
    Debug.Assert Result.Should.Contain(10)  'true

    Result.Meta.TestName = "Test - Result should end with 9 - "
    Debug.Assert Result.Should.EndWith(9)  'false
    
    Result.Meta.TestName = "Test -  - "
    Debug.Assert Result.ShouldNot.StartWith(1)  'false
    
    Result.Meta.TestName = "Test - Result shoudl not end with 0  - "
    Debug.Assert Result.ShouldNot.EndWith(0)  'false
    
    Result.Meta.TestName = "Test - result should not have length of 0 - "
    Debug.Assert Result.ShouldNot.Have.LengthOf(0)  'true
    
    Result.Meta.TestName = "Test - result should not have max length of 0 - "
    Debug.Assert Result.ShouldNot.Have.MaxLengthOf(0)  'true
    
    Result.Meta.TestName = "Test - result should not have min length of 3 - "
    Debug.Assert Result.ShouldNot.Have.MinLengthOf(3)  'true
    
    Result.Meta.TestName = "Test - result should have length of 0 - "
    Debug.Assert Result.Should.Have.LengthOf(0)  'false
    
    Result.Meta.TestName = "Test - result should have max length of 1 - "
    Debug.Assert Result.Should.Have.MaxLengthOf(1)  'false
    
    Result.Meta.TestName = "Test - result should have min length of 3 - "
    Debug.Assert Result.Should.Have.MinLengthOf(3)  'false
    
End Sub

Private Sub MetaTests(fluent As cFluentOf, testFluent As cFluent)
    Dim testResult As Boolean

    testFluent.TestValue = True
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With

    testFluent.TestValue = True
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With

    testFluent.TestValue = False
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With

    testFluent.TestValue = False
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With

    testFluent.TestValue = -1 '// -1 = true in boolean enum
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With
    
    testFluent.TestValue = -1 '// -1 = true in boolean enum
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With

    testFluent.TestValue = 0 '// 0 = false in boolean enum
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With

    testFluent.TestValue = 0 '// 0 = false in boolean enum
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With
    
    '//Approximate equality tests
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "TRUE"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "TRUE"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With

    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "FALSE"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "FALSE"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With
    
        testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "true"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "true"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With

    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "false"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(True)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(False)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(True)
    End With
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "false"
    testFluent.TestValue = testFluent.Should.Be.EqualTo(False)
    With testFluent
        Debug.Assert fluent.Of(.TestValue).Should.Be.EqualTo(True)
        Debug.Assert fluent.Of(.TestValue).ShouldNot.Be.EqualTo(False)
    End With
    
End Sub

Private Sub positiveDocumentationTests(testFluent As cFluent)
    Dim fluent As cFluent
    Dim testResult As Boolean
    Dim col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    
    Set fluent = New cFluent
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.GreaterThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.LessThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.LessThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.GreaterThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Contain(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Contain(0)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Contain(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Contain(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.StartWith(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.StartWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.EndWith(0)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.EndWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Have.LengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Have.LengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Have.MaxLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Have.MaxLengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Have.MinLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.GreaterThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.GreaterThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.GreaterThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.LessThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.LessThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.LessThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.Between(9, 11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Have.LengthBetween(1, 3)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.OneOf(9, 10, 11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ' //Object and data structure tests
    
    Set col = New Collection
    Set d = CreateObject("Scripting.Dictionary")
    Set testFluent.TestValue = col
    fluent.TestValue = testFluent.Should.Be.OneOf(col, d)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.OneOf(col, d, 10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set testFluent.TestValue = col
    fluent.TestValue = testFluent.Should.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = Nothing
    Set testFluent.TestValue = col
    fluent.TestValue = testFluent.Should.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    arr = Array(9, 10, 11)
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ReDim arr(1, 1)
    arr(0, 0) = 9
    arr(0, 1) = 10
    arr(1, 0) = 11
    arr(1, 1) = 12
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(arr)
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
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    arr = Array(9, Array(10, Array(11)))
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = New Collection
    col.Add 9
    col.Add 10
    col.Add 11
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = New Collection
    col.Add 9
    col.Add Array(10, Array(11))
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    Set col = Nothing
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, Array(10, Array(11))
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 9, 1
    d.Add 10, 2
    d.Add 11, 3
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(d.keys)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set al = CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.Should.Be.InDataStructure(al)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    ' //Approximate equality tests
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "10"
    fluent.TestValue = testFluent.Should.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "True"
    fluent.TestValue = testFluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    '//default epsilon for double comparisons is 0.000001
    '//the default can be modified by setting a value
    '//for the epsilon property in the Meta object.
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = 5.0000001
    fluent.TestValue = testFluent.Should.Be.EqualTo(5)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
End Sub

Sub negativeDocumentationTests(testFluent As cFluent)
    Dim fluent As cFluent
    Dim testResult As Boolean
    Dim col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    
    Set fluent = New cFluent
    
    testFluent.TestValue = True
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = True
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    testFluent.TestValue = False
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = False
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.GreaterThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.LessThan(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.LessThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.GreaterThan(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Contain(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Contain(0)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Contain(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Contain(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.StartWith(1)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.StartWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.EndWith(0)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.EndWith(2)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Have.LengthOf(2)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Have.LengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Have.MaxLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Have.MaxLengthOf(1)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Have.MinLengthOf(3)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.GreaterThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.GreaterThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.GreaterThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.LessThanOrEqualTo(9)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.LessThanOrEqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.LessThanOrEqualTo(11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.Between(9, 11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Have.LengthBetween(1, 3)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.OneOf(9, 10, 11)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ' //Object and data structure tests
    
    Set col = New Collection
    Set d = CreateObject("Scripting.Dictionary")
    Set testFluent.TestValue = col
    fluent.TestValue = testFluent.ShouldNot.Be.OneOf(col, d)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = 5
    fluent.TestValue = testFluent.Should.Be.OneOf(col, d, 5)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    Set col = New Collection
    Set testFluent.TestValue = col
    fluent.TestValue = testFluent.ShouldNot.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set col = Nothing
    Set testFluent.TestValue = col
    fluent.TestValue = testFluent.ShouldNot.Be.Something
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    
    arr = Array(9, 10, 11)
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ReDim arr(1, 1)
    arr(0, 0) = 9
    arr(0, 1) = 10
    arr(1, 0) = 11
    arr(1, 1) = 12
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(arr)
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
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    arr = Array(9, Array(10, Array(11)))
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(arr)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set col = New Collection
    col.Add 9
    col.Add 10
    col.Add 11
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set col = New Collection
    col.Add 9
    col.Add Array(10, Array(11))
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(col)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    Set col = Nothing
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, 10
    d.Add 3, 11
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 1, 9
    d.Add 2, Array(10, Array(11))
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(d)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add 9, 1
    d.Add 10, 2
    d.Add 11, 3
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(d.keys)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    Set al = CreateObject("System.Collections.Arraylist")
    al.Add 9
    al.Add 10
    al.Add 11
    testFluent.TestValue = 10
    fluent.TestValue = testFluent.ShouldNot.Be.InDataStructure(al)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    ' //Approximate equality tests
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "10"
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(10)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = "True"
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    '//default epsilon value for double comparisons is
    '//0.000001 the default can be modified by setting
    '//a value for the epsilon property in the Meta object.
    
    testFluent.Meta.ApproximateEqual = True
    testFluent.TestValue = 5.0000001
    fluent.TestValue = testFluent.ShouldNot.Be.EqualTo(5)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
End Sub
