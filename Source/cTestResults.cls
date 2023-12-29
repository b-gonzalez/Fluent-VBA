VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event TestPassed(ByVal TestResult As cTestResult)
Public Event TestFailed(ByVal TestResult As cTestResult)
Public Event TestCompleted(ByVal TestResult As cTestResult)

Private pResult As Boolean
Private pCol As Collection

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Public Property Get Result() As Boolean
    Result = pResult
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Property

Public Sub CheckTest(TestResult As cTestResult)
    pResult = TestResult.Result
    
    pCol.Add TestResult
    
    RaiseEvent TestCompleted(TestResult)
    
    If TestResult.Result Then
        RaiseEvent TestPassed(TestResult)
    Else
        RaiseEvent TestFailed(TestResult)
    End If
End Sub

