VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zUDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents TestResult As cTestResult
Attribute TestResult.VB_VarHelpID = -1

Sub DoWork()
    Dim fluent As cFluent
    
    Set fluent = New cFluent
    Set TestResult = fluent.Meta.TestResult
    
    fluent.TestValue = 10
    fluent.Should.Be.EqualTo (10)
    fluent.ShouldNot.Be.EqualTo (11)
End Sub

Private Sub TestResult_TestCompleted(TestDict As Object)
    IterateTestResult TestDict
End Sub

Private Sub testResult_TestFailed(TestDict As Object)
    IterateTestResult TestDict
End Sub

Private Sub testResult_TestPassed(TestDict As Object)
    IterateTestResult TestDict
End Sub

Private Sub IterateTestResult(d As Object)
    Dim elem As Variant
    
    For Each elem In d
        Debug.Print elem & ": " & d(elem)
    Next elem
    
    Debug.Print
End Sub
