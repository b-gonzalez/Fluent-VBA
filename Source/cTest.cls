VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pTestingValue As Variant
Private pTestingInput As Variant
Private pFluentPath As String
Private pResult As Boolean
Private pFunctionName As String

Public Property Let TestingValue(value As Variant)
    pTestingValue = value
End Property

Public Property Get TestingValue() As Variant
    TestingValue = pTestingValue
End Property

Public Property Let TestingInput(value As Variant)
    pTestingInput = value
End Property

Public Property Get TestingInput() As Variant
    TestingInput = pTestingInput
End Property

Public Property Let FluentPath(value As String)
    pFluentPath = value
End Property

Public Property Get FluentPath() As String
    FluentPath = pFluentPath
End Property

Public Property Let Result(value As Boolean)
    pResult = value
End Property

Public Property Get Result() As Boolean
    Result = pResult
End Property

Public Property Let FunctionName(value As String)
    pFunctionName = value
End Property

Public Property Get FunctionName() As String
    FunctionName = pFunctionName
End Property
