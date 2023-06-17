VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event TestPassed(ByVal TestDict As Scripting.Dictionary)
Public Event TestFailed(ByVal TestDict As Scripting.Dictionary)
Public Event TestCompleted(ByVal TestDict As Scripting.Dictionary)

Private pResult As Boolean

Public Property Get Result() As Boolean
    Result = pResult
End Property

Public Sub CheckTest(ByVal TestResult As Boolean, ByVal TestDict As Scripting.Dictionary)
    pResult = TestResult
    
    RaiseEvent TestCompleted(TestDict)
    
    If TestResult Then
        RaiseEvent TestPassed(TestDict)
    Else
        RaiseEvent TestFailed(TestDict)
    End If
End Sub
