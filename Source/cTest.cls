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
Private pLowerVal As Double
Private pHigherVal As Double
Private pStrTestValue As String
Private pStrTestInput As String

Public Property Let testingValue(value As Variant)
    pTestingValue = value
End Property

Public Property Set testingValue(value As Variant)
    Set pTestingValue = value
End Property

Public Property Get testingValue() As Variant
    If TypeOf pTestingValue Is Object  Then
        Set testingValue = pTestingValue
    Else
        testingValue = pTestingValue
    End If
End Property

Public Property Let testingInput(value As Variant)
    pTestingInput = value
End Property

Public Property Set testingInput(value As Variant)
    Set pTestingInput = value
End Property

Public Property Get testingInput() As Variant
    If TypeOf pTestingInput Is Object  Then
        Set testingInput = pTestingInput
    Else
        testingInput = pTestingInput
    End If
End Property

Public Property Let LowerVal(value As Double)
    pLowerVal = value
End Property

Public Property Get LowerVal() As Double
    LowerVal = pLowerVal
End Property

Public Property Let HigherVal(value As Double)
    pHigherVal = value
End Property

Public Property Get HigherVal() As Double
    HigherVal = pHigherVal
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

Public Property Let functionName(value As String)
    pFunctionName = value
End Property

Public Property Get functionName() As String
    functionName = pFunctionName
End Property

Public Property Let StrTestValue(value As String)
    pStrTestValue = value
End Property

Public Property Get StrTestValue() As String
    StrTestValue = pStrTestValue
End Property

Public Property Let StrTestInput(value As String)
    pStrTestInput = value
End Property

Public Property Get StrTestInput() As String
    StrTestInput = pStrTestInput
End Property

