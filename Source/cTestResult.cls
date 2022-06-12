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

Public Event TestPassed(fluentPath As String)
Public Event TestFailed(fluentPath As String)

Public Sub CheckTest(fluentBool As Boolean, fluentPath As String)
    If fluentBool Then
        RaiseEvent TestPassed(fluentPath)
    Else
        RaiseEvent TestFailed(fluentPath)
    End If
End Sub
