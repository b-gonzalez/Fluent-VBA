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

Public Sub CheckTest(ByVal Result As Boolean, ByVal TestDict As Scripting.Dictionary)
    RaiseEvent TestCompleted(TestDict)
    
    If Result Then
        RaiseEvent TestPassed(TestDict)
    Else
        RaiseEvent TestFailed(TestDict)
    End If
End Sub
