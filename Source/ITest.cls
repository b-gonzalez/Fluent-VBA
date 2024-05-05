VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let testingValue(ByVal value As Variant)
End Property

Public Property Set testingValue(ByVal value As Variant)
End Property

Public Property Get testingValue() As Variant
End Property

Public Property Let testingInput(ByVal value As Variant)
End Property

Public Property Set testingInput(ByVal value As Variant)
End Property

Public Property Get testingInput() As Variant
End Property

Public Property Let LowerVal(ByVal value As Double)
End Property

Public Property Get LowerVal() As Double
End Property

Public Property Let HigherVal(ByVal value As Double)
End Property

Public Property Get HigherVal() As Double
End Property

Public Property Let FluentPath(ByVal value As String)
End Property

Public Property Get FluentPath() As String
End Property

Public Property Let Result(ByVal value As Variant)
End Property

Public Property Get Result() As Variant
End Property

Public Property Let functionName(ByVal value As String)
End Property

Public Property Get functionName() As String
End Property

Public Property Let StrTestValue(ByVal value As String)
End Property

Public Property Get StrTestValue() As String
End Property

Public Property Let StrTestInput(ByVal value As String)
End Property

Public Property Get StrTestInput() As String
End Property

Public Property Let HasNull(ByVal value As Boolean)
End Property

Public Property Get HasNull() As Boolean
End Property

Public Property Let Category(ByVal value As String)
End Property

Public Property Get Category() As String
End Property

Public Property Let TestName(ByVal value As String)
End Property

Public Property Get TestName() As String
End Property
