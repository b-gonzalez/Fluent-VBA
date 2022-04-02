Attribute VB_Name = "mTests"
Option Explicit

Private Sub FluentOf()
    Dim result As IFluentOf
    Dim result2 As IFluent
    Dim TestValue As Variant
    
    Set result = New cFluent
    Set result2 = result
    TestValue = result.Of(True).Should.Be.EqualTo(True)
    result.Meta.PrintResults = True
    result2.Meta.PrintResults = True
    result2.TestValue = True
    
    result.Meta.ApproximateEqual = True
    
    With result.Of(True).Should.Be
        Debug.Assert .EqualTo("True")
    End With
'    Debug.Assert result2.Should.Be.EqualTo(True)
'    Debug.Assert result.Of(TestValue).Should.Be.EqualTo(False)
End Sub

Public Sub runMainTests()
    Dim fluent As IFluent
    Dim testFluent As IFluent
    
    Set fluent = New cFluent
    Set testFluent = New cFluent
    
    fluent.Meta.PrintTotalTests = True
    
    Call MetaTests(fluent, testFluent)
    Call positiveDocumentationTests(fluent, testFluent)
    Call negativeDocumentationTests(fluent, testFluent)
    Debug.Print "All tests Finished!"
End Sub

Private Sub Example1()
    Dim result As IFluent
    Set result = New cFluent
    result.TestValue = 10
       
    result.Meta.PrintResults = True
    
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
    
End Sub

Private Sub Example2()
    Dim testNums As Long
    Dim result As IFluent
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
    Dim result As IFluent
    Dim TestNames() As String
    Dim i As Long
    'Dim testResults(4) As Boolean
    Dim temp As Boolean
    
    Set result = New cFluent
    result.TestValue = 10
    
    With result
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
    Dim result() As IFluent
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
        result(i).Meta.TestName = TestNames(i)
        result(i).Meta.PrintResults = True
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
    Dim result() As IFluent
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
        result(i).Meta.TestName = TestNames(i)
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
        Debug.Assert temp
        Debug.Print temp
    Next i
End Sub

Private Sub Example6()
    Dim testNums As Long
    Dim result As IFluent
    Dim TestNames() As String
    Dim i As Long
    'Dim testResults(4) As Boolean
    Dim temp As Boolean
    
    Set result = New cFluent
    result.TestValue = 10
    
    result.Meta.TestName = "Test - Result should be equal to 10 - "
    Debug.Assert result.Should.Be.EqualTo(10)  ' true
    
    result.Meta.TestName = "Test - Result should greater than 9 - "
    Debug.Assert result.Should.Be.GreaterThan(9)  'true
    
    result.Meta.TestName = "Test - Result should be less than 11 - "
    Debug.Assert result.Should.Be.LessThan(11)  ' true
    
    result.Meta.TestName = "Test - Result should not be equal to 9 - "
    Debug.Assert result.ShouldNot.Be.EqualTo(9)   'true
    
    result.Meta.TestName = "Test - Result should not contain 4 - "
    Debug.Assert result.ShouldNot.Contain(4)  'true
    
    result.Meta.TestName = "Test - Result should start with 1 - "
    Debug.Assert result.Should.StartWith(1)  'true
    
    result.Meta.TestName = "Test - Result should end with 0 - "
    Debug.Assert result.Should.EndWith(0)  'true

    result.Meta.TestName = "Test - Result should contain 10 - "
    Debug.Assert result.Should.Contain(10)  'true

    result.Meta.TestName = "Test - Result should end with 9 - "
    Debug.Assert result.Should.EndWith(9)  'false
    
    result.Meta.TestName = "Test -  - "
    Debug.Assert result.ShouldNot.StartWith(1)  'false
    
    result.Meta.TestName = "Test - Result shoudl not end with 0  - "
    Debug.Assert result.ShouldNot.EndWith(0)  'false
    
    result.Meta.TestName = "Test - result should not have length of 0 - "
    Debug.Assert result.ShouldNot.Have.LengthOf(0)  'true
    
    result.Meta.TestName = "Test - result should not have max length of 0 - "
    Debug.Assert result.ShouldNot.Have.MaxLengthOf(0)  'true
    
    result.Meta.TestName = "Test - result should not have min length of 3 - "
    Debug.Assert result.ShouldNot.Have.MinLengthOf(3)  'true
    
    result.Meta.TestName = "Test - result should have length of 0 - "
    Debug.Assert result.Should.Have.LengthOf(0)  'false
    
    result.Meta.TestName = "Test - result should have max length of 1 - "
    Debug.Assert result.Should.Have.MaxLengthOf(1)  'false
    
    result.Meta.TestName = "Test - result should have min length of 3 - "
    Debug.Assert result.Should.Have.MinLengthOf(3)  'false
    
End Sub

Private Sub MetaTests(fluent As IFluent, testFluent As IFluent)
'    Dim fluent As IFluent
'    Dim testFluent As IFluent
    Dim testResult As Boolean
    
'    Set testFluent = New cFluent
'    Set fluent = New cFluent
    
    testFluent.TestValue = True
    fluent.TestValue = testFluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
    testFluent.TestValue = True
    fluent.TestValue = testFluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)

    testFluent.TestValue = False
    fluent.TestValue = testFluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
    
    testFluent.TestValue = False
    fluent.TestValue = testFluent.Should.Be.EqualTo(False)
    Debug.Assert fluent.Should.Be.EqualTo(True)
    Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
    
End Sub

Private Sub positiveDocumentationTests(fluent As IFluent, testFluent As IFluent)
    Dim testResult As Boolean
    Dim col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    
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
End Sub

Sub negativeDocumentationTests(fluent As IFluent, testFluent As IFluent)
    Dim testResult As Boolean
    Dim col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    
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
End Sub
