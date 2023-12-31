Attribute VB_Name = "mTests"
Option Explicit

Public Sub runMainTests()
    Dim fluent As cFluent
    Dim testFluent As cFluentOf
    Dim testFluentResult As cFluentOf
    Dim events As zUdeTests
    
    Set fluent = New cFluent
    Set testFluent = New cFluentOf
    Set testFluentResult = New cFluentOf
    Set events = New zUdeTests
    
    Set events.setFluent = fluent
    Set events.setFluentOf = testFluent
    Set events.setFluentEventOfResult = testFluentResult
    
    With fluent.Meta.Printing
        .TestName = "Result"
        .PassedMessage = "Success"
        .FailedMessage = "Failure"
    End With
    
'    With testFluent.Meta.Printing
'        .PassedMessage = "Success"
'        .FailedMessage = "Failure"
'    End With

    fluent.Meta.Printing.Category = "Fluent - EqualityTests"
    testFluent.Meta.Printing.Category = "Test Fluent - EqualityTests"
    Call EqualityTests(fluent, testFluent, testFluentResult)
'
    fluent.Meta.Printing.Category = "Fluent - positiveDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - positiveDocumentationTests"
    Call positiveDocumentationTests(fluent, testFluent, testFluentResult)

    fluent.Meta.Printing.Category = "Fluent - negativeDocumentationTests"
    testFluent.Meta.Printing.Category = "Test Fluent - negativeDocumentationTests"
    Call negativeDocumentationTests(fluent, testFluent, testFluentResult)
'
'    Debug.Print "All tests Finished!"
    Call printTestCount(testFluent.Meta.Tests.Count)
    
'    fluent.Meta.Printing.PrintToSheet
'    testFluent.Meta.Printing.PrintToSheet
'    fluent.Meta.Printing.PrintToImmediate
End Sub

Private Sub printTestCount(TestCount As Long)
    If TestCount > 1 Then
        Debug.Print TestCount & " tests finished!"
    ElseIf TestCount = 1 Then
        Debug.Print "Test finished!"
    End If
End Sub

Private Sub EqualityTests(fluent As cFluent, testFluent As cFluentOf, testFluentResult As cFluentOf)
    Dim test As cTest
    Dim i As Long

    With fluent.Meta.Tests
        fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(False).Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        fluent.TestValue = testFluent.Of(True).Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(-1).Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        fluent.TestValue = testFluent.Of(0).Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        '//Approximate equality tests
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("TRUE").Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("FALSE").Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("true").Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("false").Should.Be.EqualTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.Result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).Result
    Next i
End Sub

Private Sub positiveDocumentationTests(fluent As cFluent, testFluent As cFluentOf, testFluentResult As cFluentOf)
    'Dim testFluent As cFluentOf
    Dim test As cTest
    Dim Col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    Dim i As Long
    
    'Set testFluent = New cFluentOf
    With fluent.Meta.Tests
        fluent.TestValue = testFluent.Of(10).Should.Be.EqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(9)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(9)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThan(11)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThan(11)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(1)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(0)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Contain(2)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.StartWith(1)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        fluent.TestValue = testFluent.Of(10).Should.StartWith(2)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.EndWith(0)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        
        fluent.TestValue = testFluent.Of(10).Should.EndWith(2)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(2)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthOf(1)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(3)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MaxLengthOf(1)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Have.MinLengthOf(3)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(9)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    
        fluent.TestValue = testFluent.Of(10).Should.Be.GreaterThanOrEqualTo(11)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(9)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.LessThanOrEqualTo(11)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.Between(9, 11)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Have.LengthBetween(1, 3)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(9, 10, 11)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ' //Object and data structure tests
        
        Set Col = New Collection
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(Col).Should.Be.OneOf(Col, d)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        fluent.TestValue = testFluent.Of(10).Should.Be.OneOf(Col, d, 10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(Col).Should.Be.Something
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = Nothing
        fluent.TestValue = testFluent.Of(Col).Should.Be.Something
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array()
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = Nothing
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
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
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add 10
        Col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set Col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.keys)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        'with explicit flRecursive
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
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
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add 10
        Col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set Col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.keys, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al, flRecursive)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ' //with explicit flIterative
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
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
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(arr, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add 10
        Col.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(Col, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set Col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(d.keys, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).Should.Be.InDataStructure(al, flIterative)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        
        ' //Approximate equality tests
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("10").Should.Be.EqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("True").Should.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        '//default epsilon for double comparisons is 0.000001
        '//the default can be modified by setting a value
        '//for the epsilon property in the Meta object.
        
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of(5.0000001).Should.Be.EqualTo(5)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluent.Meta.ApproximateEqual = False
        
        '//Evaluation tests
        
        fluent.TestValue = testFluent.Of(True).Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(True).Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(False).Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(False).Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("true").Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("false").Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("TRUE").Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("FALSE").Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(-1).Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(-1).Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(0).Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(0).Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("-1").Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("-1").Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("0").Should.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("0").Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(5 + 5).Should.EvaluateTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("5 + 5").Should.EvaluateTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("5 + 5 = 10").Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("5 + 5 > 9").Should.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        '//Testing errors is possible if they're put in strings
        fluent.TestValue = testFluent.Of("1 / 0").Should.EvaluateTo(CVErr(xlErrDiv0))
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        
        fluent.TestValue = testFluent.Of("1 / 0").Should.EvaluateTo(CVErr(xlErrDiv0))
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("abc").Should.Be.Alphabetic
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(123).Should.Be.Numeric
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("abc123").Should.Be.Alphanumeric
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
    End With
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.Result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).Result
    Next i
    
End Sub

Private Sub negativeDocumentationTests(fluent As cFluent, testFluent As cFluentOf, testFluentResult As cFluentOf)
    'Dim testFluent As cFluentOf
    Dim test As cTest
    Dim Col As Collection
    Dim arr As Variant
    Dim d As Object
    Dim al As Object
    Dim i As Long
    
    'Set testFluent = New cFluentOf
    With fluent.Meta.Tests
    
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.EqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(9)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(9)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThan(11)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThan(11)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(1)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(0)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Contain(2)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(1)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.StartWith(2)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(0)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.EndWith(2)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(2)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthOf(1)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(3)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MaxLengthOf(1)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.MinLengthOf(3)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(9)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.GreaterThanOrEqualTo(11)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(9)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.LessThanOrEqualTo(11)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.Between(9, 11)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Have.LengthBetween(1, 3)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(9, 10, 11)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ' //Object and data structure tests
        
        Set Col = New Collection
        Set d = New Scripting.Dictionary
        fluent.TestValue = testFluent.Of(Col).ShouldNot.Be.OneOf(Col, d)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.OneOf(Col, d, 10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(Col).ShouldNot.Be.Something
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = Nothing
        fluent.TestValue = testFluent.Of(Col).ShouldNot.Be.Something
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array()
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = Nothing
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(123)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array(9, 10, 11)
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ReDim arr(1, 1)
        arr(0, 0) = 9
        arr(0, 1) = 10
        arr(1, 0) = 11
        arr(1, 1) = 12
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
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
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        arr = Array(9, Array(10, Array(11)))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(arr)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add 10
        Col.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(Col)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        Set Col = New Collection
        Col.Add 9
        Col.Add Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(Col)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set Col = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, 10
        d.Add 3, 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 1, 9
        d.Add 2, Array(10, Array(11))
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set d = New Scripting.Dictionary
        d.Add 9, 1
        d.Add 10, 2
        d.Add 11, 3
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(d.keys)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        Set d = Nothing
        
        Set al = CreateObject("System.Collections.Arraylist")
        al.Add 9
        al.Add 10
        al.Add 11
        fluent.TestValue = testFluent.Of(10).ShouldNot.Be.InDataStructure(al)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        ' //Approximate equality tests
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("10").ShouldNot.Be.EqualTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of("True").ShouldNot.Be.EqualTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        '//default epsilon for double comparisons is 0.000001
        '//the default can be modified by setting a value
        '//for the epsilon property in the Meta object.
        
        testFluent.Meta.ApproximateEqual = True
        fluent.TestValue = testFluent.Of(5.0000001).ShouldNot.Be.EqualTo(5)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluent.Meta.ApproximateEqual = False
        
        '//Evaluation tests
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(True).ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(False).ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("true").ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("false").ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("TRUE").ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("FALSE").ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(-1).ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(-1).ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(0).ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(0).ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of("-1").ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("-1").ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of("0").ShouldNot.EvaluateTo(False)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("0").ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(True)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(False)
        
        fluent.TestValue = testFluent.Of(5 + 5).ShouldNot.EvaluateTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("5 + 5").ShouldNot.EvaluateTo(10)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("5 + 5 = 10").ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("5 + 5 > 9").ShouldNot.EvaluateTo(True)
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        '//Testing errors is possible if they're put in strings
        fluent.TestValue = testFluent.Of("1 / 0").ShouldNot.EvaluateTo(CVErr(xlErrDiv0))
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("abc123").ShouldNot.Be.Alphanumeric
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of(123).ShouldNot.Be.Numeric
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        
        fluent.TestValue = testFluent.Of("abc123").ShouldNot.Be.Alphanumeric
        Debug.Assert fluent.Should.Be.EqualTo(False)
        Debug.Assert fluent.ShouldNot.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).Should.Be.EqualTo(True)
        Debug.Assert testFluentResult.Of(.Result).ShouldNot.Be.EqualTo(False)
        testFluentResult.Of(.Result).Should.Be.EqualTo (False) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
        testFluentResult.Of(.Result).ShouldNot.Be.EqualTo (True) '//Not asserting. Intentionally failing to test TestFailed event linked to object.
    End With
    
    For Each test In fluent.Meta.Tests
        Debug.Assert test.Result
    Next test
    
    For i = 1 To fluent.Meta.Tests.Count
        Debug.Assert fluent.Meta.Tests(i).Result
    Next i
End Sub
