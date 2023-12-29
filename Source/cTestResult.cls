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

Private pResult As Boolean
Private pFluentPath As String
Private pExpectedValue As Variant
Private pActualValue As Variant

Public Property Let Result(value As Boolean)
    pResult = value
End Property

Public Property Get Result() As Boolean
    Result = pResult
End Property

Public Property Let FluentPath(value As String)
    pFluentPath = value
End Property

Public Property Get FluentPath() As String
    FluentPath = pFluentPath
End Property

Public Property Let ExpectedValue(value As Variant)
    pExpectedValue = value
End Property

Public Property Get ExpectedValue() As Variant
    ExpectedValue = pExpectedValue
End Property

Public Property Let ActualValue(value As Variant)
    pActualValue = value
End Property

Public Property Get ActualValue() As Variant
    ActualValue = pActualValue
End Property
