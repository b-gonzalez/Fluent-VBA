VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let Algorithm(ByVal value As Variant)
End Property

Public Property Get Algorithm() As Variant
End Property

Public Property Let AlgorithmValueSet(ByVal value As Boolean)
End Property

Public Property Get AlgorithmValueSet() As Boolean
End Property

Public Property Let negateValue(ByVal value As Boolean)
End Property

Public Property Get negateValue() As Boolean
End Property

Public Property Let ResultNumber(ByVal value As Long)
End Property

Public Property Get ResultNumber() As Long
End Property

Public Property Let TestValueRecur(ByVal value As String)
End Property

Public Property Get TestValueRecur() As String
End Property

Public Property Let TestValueIter(ByVal value As String)
End Property

Public Property Get TestValueIter() As String
End Property

Public Property Let TestInputRecur(ByVal value As String)
End Property

Public Property Get TestInputRecur() As String
End Property

Public Property Let TestInputIter(ByVal value As String)
End Property

Public Property Get TestInputIter() As String
End Property

Public Property Let IsRecurIterFunc(ByVal value As Boolean)
End Property

Public Property Get IsRecurIterFunc() As Boolean
End Property

Public Property Let IsBaseCaseRecur(ByVal value As Boolean)
End Property

Public Property Get IsBaseCaseRecur() As Boolean
End Property

Public Property Let IsBaseCaseIter(ByVal value As Boolean)
End Property

Public Property Get IsBaseCaseIter() As Boolean
End Property
