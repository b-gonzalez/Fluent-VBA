Attribute VB_Name = "mTests"
Option Explicit

Private mCounter As Long
Private mTestCounter As Long

Public Enum hw
    helloWorld
    goodbyeWorld
End Enum

Public Sub runMainTests()
    Dim fluent As IFluent
    Dim testFluent As IFluentOf
    Dim testFluentResult As IFluentOf
    Dim events As zUdeTests
    Dim posTestFluent As IFluentOf
    Dim negTestFluent As IFluentOf
    Dim nulTestFluent As IFluentOf
    Dim col As Collection
    Dim i As Long
    Dim posFluentOfStr As String
    Dim negFluentOfStr As String
    Dim tempCounter As Long
    Dim tempDict As Scripting.Dictionary
    
    Set fluent = New cFluent
    Set testFluent = New cFluentOf
    Set testFluentResult = New cFluentOf
    Set events = New zUdeTests
    
    Set events.setFluent = fluent
    Set events.setFluentOf = testFluent
    Set events.setFluentEventOfResult = testFluentResult
    
'    With fluent.Meta.Printing
'        .TestName = "Result"
'        .PassedMessage = "Success"
'        .FailedMessage = "Failure"
'    End With
    
    With testFluent.Meta.Printing
        .PassedMessage = "Success"
        .FailedMessage = "Failure"
        .UnexpectedMessage = "What?"
    End With

    fluent.Meta.Printing.Category = "Fluent - EqualityTests"
    testFluent.Meta.Printing.Category = "Test Fluent - EqualityTests"
    Call EqualityTests(fluent, testFluent, testFluentResult)
'
    fluent.Meta.Printing.Category = "Fluent - positiveDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - positiveDocumentationTests"
    Set posTestFluent = positiveDocumentationTests(fluent, testFluent, testFluentResult)
    Debug.Assert validateTestDictCounters(testFluent.Meta.Tests.TestDictCounter)
    
'    posFluentOfStr = getFluentOfCounts(posTestFluent)
    
    fluent.Meta.Printing.Category = "Fluent - negativeDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - negativeDocumentationTests"
    Set negTestFluent = negativeDocumentationTests(fluent, testFluent, testFluentResult)
    Debug.Assert validateTestDictCounters(testFluent.Meta.Tests.TestDictCounter)
    
'    negFluentOfStr = getFluentOfCounts(negTestFluent)
    
    With posTestFluent.Meta
        For i = 1 To .Tests.Count
            Debug.Assert .Tests(i).functionName = negTestFluent.Meta.Tests(i).functionName
        Next i
    End With
    
    Debug.Assert validateNegativeCounters(testFluent)

    tempCounter = mCounter

    Set fluent = New cFluent
    Set testFluent = New cFluentOf

    mCounter = 0
    fluent.Meta.Printing.Category = "Fluent - nullDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - nullDocumentationTests"
    Set nulTestFluent = nullDocumentationTests(fluent, testFluent, testFluentResult)
    Set tempDict = testFluent.Meta.Tests.TestDictCounter
    tempDict("OneOf") = 1 '//intentionally passing since this method cannot be checked for nulls
    Debug.Assert validateTestDictCounters(tempDict, counter:=1) '// set to 1 to account for intentionally passed OneOf method.
    
    mCounter = tempCounter
    
    mCounter = mCounter + nulTestFluent.Meta.Tests.Count
    
    Set fluent = New cFluent
    
    mCounter = mCounter + MiscTests(fluent)

    Debug.Print "All tests Finished"
    Call printTestCount(mCounter)
    
    mCounter = 0
    
    mTestCounter = 0

    Debug.Assert events.CheckTestCounters

    Debug.Assert checkResetCounters(fluent, testFluent)
    
'    testFluent.Meta.Printing.PrintToSheet
    
End Sub

Private Sub printTestCount(testCount As Long)
    If testCount > 1 Then
        Debug.Print testCount & " tests finished!" & vbNewLine
    ElseIf testCount = 1 Then
        Debug.Print "1 Test finished!"
    End If
End Sub

Private Sub TrueAssertAndRaiseEvents(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.Tests.Count = mCounter

    With fluent.Meta.Tests
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
End Sub

Private Sub FalseAssertAndRaiseEvents(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    mCounter = mCounter + 1
    mTestCounter = mTestCounter + 1
    
    Debug.Assert testFluent.Meta.Tests.Count = mCounter

    With fluent.Meta.Tests
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
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

Private Sub EqualityTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
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
    
        fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Approximate equality tests
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(False)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(False)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.ApproximateEqual = False
        
        '//Null and Empty tests
        
        fluent.TestValue = testFluent.Of(Null).Should.Be.EqualTo(Null)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(Null).ShouldNot.Be.EqualTo(Null)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of().Should.Be.EqualTo(Empty)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of().ShouldNot.Be.EqualTo(Empty)
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
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            resultBool = test.result = .Tests(i).result
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
End Sub

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
    
    With fluent.Meta.Tests
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
        
        fluent.TestValue = testFluent.Of(10).Should.StartWith(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).Should.StartWith(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.EndWith(0)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.EndWith(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").Should.Have.LengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MinLengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(9, 11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(1, 3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(9, 10, 11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
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
        
        arr = Array()
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'        Set col = Nothing
'        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col)
'        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
'
'        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(10)
'        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
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
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.Keys)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        'with explicit flRecursive
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
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
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.Keys, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ' //with explicit flIterative
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
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
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.Keys, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        ReDim arr(1, 1)
        arr(0, 0) = 12
        arr(0, 1) = 13
        arr(1, 0) = 14
        arr(1, 1) = 15
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(12).Should.Be.InDataStructures(flRecursive, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
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
        fluent.TestValue = testFluent.Of(9).Should.Be.InDataStructures(flRecursive, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        fluent.TestValue = testFluent.Of(13).Should.Be.InDataStructures(flRecursive, col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(16).Should.Be.InDataStructures(flRecursive, arr, col, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        arr = Array(12, 13, 14)
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(14).Should.Be.InDataStructures(flRecursive, col, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
    
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(2).Should.Be.InDataStructures(flRecursive, d.Items, d.Keys)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = Nothing
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(6, Array(7, Array(8)))
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(8).Should.Be.InDataStructures(flRecursive, al, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        ReDim arr(1, 1)
        arr(0, 0) = 12
        arr(0, 1) = 13
        arr(1, 0) = 14
        arr(1, 1) = 15
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(12).Should.Be.InDataStructures(flIterative, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
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
        fluent.TestValue = testFluent.Of(9).Should.Be.InDataStructures(flIterative, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        fluent.TestValue = testFluent.Of(13).Should.Be.InDataStructures(flIterative, col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(15, 16, 17)
        fluent.TestValue = testFluent.Of(16).Should.Be.InDataStructures(flIterative, arr, col, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, col)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        arr = Array(12, 13, 14)
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(14).Should.Be.InDataStructures(flIterative, col, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
    
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(2).Should.Be.InDataStructures(flIterative, d.Items, d.Keys)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = Nothing
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, d)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, al)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(6, Array(7, Array(8)))
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(8).Should.Be.InDataStructures(flIterative, al, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        ' //Approximate equality tests
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("10").Should.Be.EqualTo(10)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("True").Should.Be.EqualTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//default epsilon for double comparisons is 0.000001
        '//the default can be modified by setting a value
        '//for the epsilon property in the Meta object.
        
        testFluent.Meta.ApproximateEqual = True
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
        fluent.TestValue = testFluent.Of(TypeName(arr) = "Variant()").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(IsArray(arr)).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(TypeName(col) = "Collection").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf col Is Collection).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = New Dictionary
        fluent.TestValue = testFluent.Of(TypeName(d) = "Dictionary").Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf d Is Dictionary).Should.EvaluateTo(True)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Testing errors is possible if they're put in strings
        fluent.TestValue = testFluent.Of("1 / 0").Should.EvaluateTo(CVErr(xlErrDiv0))
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").Should.Be.Alphabetic
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(123).Should.Be.Numeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc123").Should.Be.Alphanumeric
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
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
        Debug.Assert test.result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            resultBool = test.result = .Tests(i).result
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput
            
            Debug.Assert resultBool And fluentBool And valueBool And inputBool
            
            i = i + 1
        Next test
    End With
    
    Debug.Print "Positive tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
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
    
    With fluent.Meta.Tests
    
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
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(1)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(0)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Have.LengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(1)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MinLengthOf(3)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(9)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(11)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(9)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(9, 11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(1, 3)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(9, 10, 11)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
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
        
        arr = Array()
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'        Set col = Nothing
'        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col)
'        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
'        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(123)
'        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
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
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d.Keys)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        
        
        'with explicit flRecursive
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
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
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.Keys, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al, flRecursive)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ' //with explicit flIterative
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
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
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add 10
        col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(col, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.Keys, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al, flIterative)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)



        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, arr)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 12
        arr(0, 1) = 13
        arr(1, 0) = 14
        arr(1, 1) = 15
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(12).Should.Be.InDataStructures(flRecursive, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
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
        fluent.TestValue = testFluent.Of(9).Should.Be.InDataStructures(flRecursive, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, arr, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        fluent.TestValue = testFluent.Of(13).ShouldNot.Be.InDataStructures(flRecursive, col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(9, 10, 16)
        fluent.TestValue = testFluent.Of(16).ShouldNot.Be.InDataStructures(flRecursive, arr, col, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        arr = Array(12, 13, 14)
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(14).ShouldNot.Be.InDataStructures(flRecursive, col, arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
    
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(2).ShouldNot.Be.InDataStructures(flRecursive, d.Items, d.Keys)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(6, Array(7, Array(8)))
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(8).ShouldNot.Be.InDataStructures(flRecursive, al, arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flRecursive, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        ReDim arr(1, 1)
        arr(0, 0) = 12
        arr(0, 1) = 13
        arr(1, 0) = 14
        arr(1, 1) = 15
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(12).Should.Be.InDataStructures(flIterative, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
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
        fluent.TestValue = testFluent.Of(9).Should.Be.InDataStructures(flIterative, arr, arr2)
        Call TrueAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, arr, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        fluent.TestValue = testFluent.Of(13).ShouldNot.Be.InDataStructures(flIterative, col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 12
        col.Add 13
        col.Add 14
        arr = Array(9, Array(10, Array(11)))
        arr2 = Array(9, 10, 16)
        fluent.TestValue = testFluent.Of(16).ShouldNot.Be.InDataStructures(flIterative, arr, col, arr2)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, col)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
        
        arr = Array(12, 13, 14)
        Set col = New Collection
        col.Add 9
        col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(14).ShouldNot.Be.InDataStructures(flIterative, col, arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set col = Nothing
    
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(2).ShouldNot.Be.InDataStructures(flIterative, d.Items, d.Keys)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, d)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        arr = Array(6, Array(7, Array(8)))
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(8).ShouldNot.Be.InDataStructures(flIterative, al, arr)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructures(flIterative, al)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
        ' //Approximate equality tests
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("10").ShouldNot.Be.EqualTo(10)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("True").ShouldNot.Be.EqualTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//default epsilon for double comparisons is 0.000001
        '//the default can be modified by setting a value
        '//for the epsilon property in the Meta object.
        
        testFluent.Meta.ApproximateEqual = True
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
        fluent.TestValue = testFluent.Of(TypeName(arr) = "Variant()").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(IsArray(arr)).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set col = New Collection
        fluent.TestValue = testFluent.Of(TypeName(col) = "Collection").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf col Is Collection).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        Set d = New Dictionary
        fluent.TestValue = testFluent.Of(TypeName(d) = "Dictionary").ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(TypeOf d Is Dictionary).ShouldNot.EvaluateTo(True)
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        '//Testing errors is possible if they're put in strings
        fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.EvaluateTo(CVErr(xlErrDiv0))
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc").ShouldNot.Be.Alphabetic
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of(123).ShouldNot.Be.Numeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
        fluent.TestValue = testFluent.Of("abc123").ShouldNot.Be.Alphanumeric
        Call FalseAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
        
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
        Debug.Assert test.result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).result
    Next i
    
    i = 1
    
    With testFluent.Meta
        For Each test In .Tests
            resultBool = test.result = .Tests(i).result
            fluentBool = test.FluentPath = .Tests(i).FluentPath
            valueBool = test.StrTestValue = .Tests(i).StrTestValue
            inputBool = test.StrTestInput = .Tests(i).StrTestInput
            
            Debug.Assert resultBool And fluentBool And valueBool And inputBool
            
            i = i + 1
        Next test
    End With
    
    Debug.Print "Negative tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set negativeDocumentationTests = testFluent
    
End Function

Public Function nullDocumentationTests(fluent As IFluent, testFluent As IFluentOf, testFluentResult As IFluentOf)
    Dim col As Collection
    Dim d As Dictionary
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

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, "Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, 123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flRecursive, Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, "Hello World")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, 10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, 123.45)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, True)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructures(flIterative, Null)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Be.GreaterThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
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
    
    fluent.TestValue = testFluent.Of(Array()).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThan(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
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
    
    fluent.TestValue = testFluent.Of(Array()).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Be.LessThanOrEqualTo(10)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)

    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Contain(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Contain(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Contain("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.StartWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.StartWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.StartWith(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.StartWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set col = New Collection
    fluent.TestValue = testFluent.Of(col).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    col.Add "Hello world!"
    fluent.TestValue = testFluent.Of(col).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array()).Should.EndWith("Hello")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.EndWith(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.EndWith("Hello world!")
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.EndWith(2.34)
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
    
    fluent.TestValue = testFluent.Of(Array()).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Array(1, 2, 3)).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    Set d = CreateObject("Scripting.Dictionary")
    fluent.TestValue = testFluent.Of(d).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    d.Add 1, 1.23
    d.Add 2, 2.34
    d.Add 3, 3.34
    fluent.TestValue = testFluent.Of(d).Should.Have.MinLengthOf(2.34)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of(Null).Should.Have.MinLengthOf(2)
    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
    fluent.TestValue = testFluent.Of("Hello World!").Should.Be.Something
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
    
'    fluent.TestValue = testFluent.Of("Hello World!").Should.Have.LengthBetween(1, 10)
'    Call NullAssertAndRaiseEvents(fluent, testFluent, testFluentResult)
    
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
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Be.IdenticalTo("Hello world")
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
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameUniqueElementsAs("Hello world")
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
    
    Set d = New Scripting.Dictionary
    fluent.TestValue = testFluent.Of(d).Should.Have.SameElementsAs("Hello world")
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
    
    Debug.Print "Null tests finished"
    printTestCount (mTestCounter)
    mTestCounter = 0
    
    Set nullDocumentationTests = testFluent
End Function

Private Function MiscTests(fluent As IFluent)
    Dim testCount As Long
    
    Debug.Assert fluent.Should.Be.EqualTo(Empty)
    
    Debug.Print "Misc tests finished"
    testCount = fluent.Meta.Tests.Count
    printTestCount (testCount)
    Debug.Print
    
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

Private Function validateTestDictCounters(d As Scripting.Dictionary, Optional counter As Long = 0)
    Dim elem As Variant
    Dim b As Boolean
    
    For Each elem In d.Keys
        If d(elem) > 0 Then
            counter = counter + 1
        End If
    Next elem
    
    validateTestDictCounters = (d.Count = counter)
End Function

Private Function validateNegativeCounters(testFluent As IFluentOf) As Boolean
    Dim d As Scripting.Dictionary
    Dim test As ITest
    Dim counter As Long
    Dim fn As String
    Dim elem As Variant
    
    Set d = New Scripting.Dictionary
    
    counter = 0
    
    For Each test In testFluent.Meta.Tests
        If test.NegateValue Then
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
