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

Public Event TestPassed(TestDict As Object)
Public Event TestFailed(TestDict As Object)

Public Sub CheckTest(result As Boolean, TestDict As Object)
    If result Then
        RaiseEvent TestPassed(TestDict)
    Else
        RaiseEvent TestFailed(TestDict)
    End If
End Sub
